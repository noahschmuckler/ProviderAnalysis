#requires -Version 5.1
# ProviderAnalysisServer.ps1 - Draft 6.0 (ranked outreach lists, patient flags, Communication tab, unattended-run recovery, report server feeds, group Overview, new-patient lists, POC coding, Combine)
# Adds provider indexes, guided cross-source mapping, fuzzy suggestions, NPI profiles, prepared job tracking, ranked outreach lists,
# a flag mode (/flag?jobId=...) that excludes flagged patients from every list and panel on every later report for that provider,
# and a Communication tab (/communication) that emails every provider in a risk pool a fresh report PDF via Outlook drafts.
# The POC Coding tab (/poc-coding) analyzes point-of-care billing and coding for one department of a billing export and saves each run as an HTML report.
# The Combine tab (/combine) fingerprints loaded .xlsx/.csv exports and combines them into one row per patient (plus summary views) under an
# editable, auditable JSON config, with a copy-paste round trip through an enterprise LLM.
[CmdletBinding()]
param([int]$PreferredPort=8765,[switch]$NoBrowser,[switch]$SkipModuleInstallPrompt,[string]$RunJobId='',[string]$RunOverviewId='',[string]$RunNewListId='',[string]$RunPocId='',[string]$RunCombineId='')
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
 Overviews=Join-Path $script:Root 'overviews'
 NewLists=Join-Path $script:Root 'new-patient-lists'
 Poc=Join-Path $script:Root 'poc-coding'
 PocFiles=Join-Path (Join-Path $script:Root 'poc-coding') 'source-files'
 Combine=Join-Path $script:Root 'combine'
 CombineFiles=Join-Path (Join-Path $script:Root 'combine') 'source-files'
 CombineConfigs=Join-Path (Join-Path $script:Root 'combine') 'configs'
 CombineHistory=Join-Path (Join-Path (Join-Path $script:Root 'combine') 'configs') 'history'
 CombineOutputs=Join-Path (Join-Path $script:Root 'combine') 'outputs'
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
function Json([string]$Path){
 # Reads a JSON record that worker processes may be rewriting at the same moment. The file is opened letting others read, write and replace
 # it, so a reader never blocks Save-JsonAtomic in another process. A moment when it is being swapped in, or held by antivirus or a sync
 # client such as OneDrive ("being used by another process"), is retried for about six seconds instead of failing the request; a record
 # whose temp file shows a save in progress is waited for rather than reported missing.
 for($attempt=1;;$attempt++){
  if(![IO.File]::Exists($Path)){
   $saving=$false;try{$dir=[IO.Path]::GetDirectoryName($Path);if($dir -and [IO.Directory]::Exists($dir)){$saving=(@([IO.Directory]::GetFiles($dir,[IO.Path]::GetFileName($Path)+'.*.tmp')).Count -gt 0)}}catch{}
   if($saving -and $attempt -lt 25){Start-Sleep -Milliseconds ([Math]::Min(400,20*$attempt));continue};return $null
  }
  try{
   $fs=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
   try{$sr=New-Object IO.StreamReader($fs,[Text.Encoding]::UTF8,$true);$text=$sr.ReadToEnd()}finally{$fs.Dispose()}
   if([string]::IsNullOrWhiteSpace($text)){if($attempt -lt 10){Start-Sleep -Milliseconds 100;continue};return $null}
   try{return ($text|ConvertFrom-Json)}catch{if($attempt -lt 10){Start-Sleep -Milliseconds 100;continue};throw ('Could not read '+$Path+': '+$_.Exception.Message)}
  }
  catch [IO.FileNotFoundException],[IO.DirectoryNotFoundException]{return $null}
  catch [IO.IOException],[UnauthorizedAccessException]{if($attempt -ge 25){throw};Start-Sleep -Milliseconds ([Math]::Min(400,20*$attempt))}
 }
}

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
function Save-JsonAtomic([string]$Path,[object]$Value){
 # Writes a per-process temp file, then swaps it in with File.Replace so readers never see the file missing. On Windows the swap fails while
 # something else has the file open without sharing it ("Cannot create a file when that file already exists", "being used by another
 # process": another process reading it, antivirus, or a sync client such as OneDrive), so it is retried for about fifteen seconds; from the
 # tenth try an overwrite in place is tried as well, which works where only the replace is refused (Json waits out a half-written file).
 $tmp=$Path+'.'+$PID+'.tmp';$Value|ConvertTo-Json -Depth 15|Set-Content -LiteralPath $tmp -Encoding UTF8
 for($attempt=1;;$attempt++){
  try{if([IO.File]::Exists($Path)){[IO.File]::Replace($tmp,$Path,[NullString]::Value)}else{[IO.File]::Move($tmp,$Path)};return}
  catch{
   if($attempt -ge 10 -and [IO.File]::Exists($Path)){try{[IO.File]::Copy($tmp,$Path,$true);Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue;return}catch{}}
   if($attempt -ge 40){Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue;throw};Start-Sleep -Milliseconds ([Math]::Min(500,25*$attempt))
  }
 }
}
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
 try{$clientDate=[string]$Context.Request.Headers['X-File-Timestamp'];if($clientDate){$dt=[DateTimeOffset]::FromUnixTimeMilliseconds([int64]$clientDate).UtcDateTime;(Get-Item $path).CreationTimeUtc=$dt}}catch{Remove-Item $path -Force -ErrorAction SilentlyContinue;throw}
 return (Register-StagedFile $token $path ([IO.Path]::GetFileName($name)) 'upload')
}
function Register-StagedFile([string]$Token,[string]$Path,[string]$OriginalName,[string]$Via){
 # Fingerprints a staged XLSX against the source registry and writes its staging record; used by browser uploads and report server downloads alike.
 $metaPath=Join-Path $script:Paths.Staging ($Token+'.json')
 try{$match=Test-WorkbookSchema $Path;$hash=Get-FileHash256 $Path;$file=Get-Item $Path;$meta=[ordered]@{token=$Token;sourceKey=$match.sourceKey;displayName=$match.displayName;canonicalFileName=$match.canonicalFileName;originalName=$OriginalName;stagedPath=$Path;sha256=$hash;byteLength=$file.Length;worksheet=$match.worksheet;headerRow=$match.headerRow;rowCount=$match.rowCount;sourceFileCreatedUtc=$file.CreationTimeUtc.ToString('o');stagedUtc=[DateTime]::UtcNow.ToString('o');uploadedBy=[Environment]::UserName;via=$Via};Save-JsonAtomic $metaPath $meta;Log 'UPLOAD_VALIDATED' 'OK' ('Identified '+$match.sourceKey+'; '+$match.rowCount+' rows ('+$Via+').') $match.sourceKey;return $meta}
 catch{Remove-Item $Path,$metaPath -Force -ErrorAction SilentlyContinue;Log 'UPLOAD_VALIDATION' 'FAILED' $_.Exception.Message;throw}
}
function Publish-StagedFile([string]$Token){
 if($Token -notmatch '^[a-f0-9]{32}$'){throw 'Invalid confirmation token.'};$metaPath=Join-Path $script:Paths.Staging ($Token+'.json');$meta=Json $metaPath;if($null -eq $meta){throw 'Staged upload was not found or expired.'};$sourceKey=[string]$meta.sourceKey;$target=Join-Path $script:Paths.CanonicalCurrent ([string]$meta.canonicalFileName);$archiveDir=Join-Path $script:Paths.CanonicalArchive $sourceKey;if(!(Test-Path $archiveDir)){New-Item -ItemType Directory -Path $archiveDir -Force|Out-Null};$archive=$null;$promoted=$false;$stamp=Get-Date -Format 'yyyyMMdd-HHmmss'
 try{if(Test-Path $target){$oldHash=Get-FileHash256 $target;$archive=Join-Path $archiveDir ($sourceKey+'__'+$stamp+'__'+$oldHash.Substring(0,8)+'.xlsx');$n=1;while(Test-Path -LiteralPath $archive){$archive=Join-Path $archiveDir ($sourceKey+'__'+$stamp+'__'+$oldHash.Substring(0,8)+'_'+$n+'.xlsx');$n++};Move-Item -LiteralPath $target -Destination $archive -ErrorAction Stop};Move-Item -LiteralPath ([string]$meta.stagedPath) -Destination $target -ErrorAction Stop;$promoted=$true;$manifest=Json $script:ManifestPath;if($null -eq $manifest){$manifest=[pscustomobject]@{manifestVersion=1;updatedUtc=$null;sources=[pscustomobject]@{}}};if($null -eq $manifest.sources){$manifest|Add-Member -NotePropertyName sources -NotePropertyValue ([pscustomobject]@{}) -Force};$entry=[ordered]@{sourceKey=$sourceKey;currentPath=$target;sha256=[string]$meta.sha256;byteLength=[int64]$meta.byteLength;rowCount=[int]$meta.rowCount;worksheet=[string]$meta.worksheet;headerRow=[int]$meta.headerRow;fileLastWriteUtc=(Get-Item $target).LastWriteTimeUtc.ToString('o');sourceFileCreatedUtc=[string]$meta.sourceFileCreatedUtc;importedUtc=[DateTime]::UtcNow.ToString('o');importedBy=[Environment]::UserName;previousArchivePath=$archive};$manifest.sources|Add-Member -NotePropertyName $sourceKey -NotePropertyValue $entry -Force;$manifest.updatedUtc=[DateTime]::UtcNow.ToString('o');Save-JsonAtomic $script:ManifestPath $manifest;Remove-Item $metaPath -Force;Log 'CANONICAL_REPLACED' 'OK' ('Archived prior file and imported '+$meta.originalName) $sourceKey;return $entry}catch{$failure=$_.Exception.Message;try{if($promoted -and (Test-Path $target)){Move-Item $target ([string]$meta.stagedPath) -Force};if($archive -and (Test-Path $archive)){Move-Item $archive $target -Force}}catch{$failure+='; rollback error: '+$_.Exception.Message};Log 'CANONICAL_REPLACE' 'FAILED' $failure $sourceKey;throw $failure}
}
function Remove-ExpiredStaging{Get-ChildItem $script:Paths.Staging -File -ErrorAction SilentlyContinue|Where-Object{$_.LastWriteTimeUtc -lt [DateTime]::UtcNow.AddDays(-2)}|Remove-Item -Force -ErrorAction SilentlyContinue}
function ConvertTo-HtmlSafe([object]$Value){if($null -eq $Value){''}else{[Net.WebUtility]::HtmlEncode([string]$Value)}}
function Page{
 $rows='';foreach($r in (Status)){$state=$(if($r.present){'Present'}else{'Missing'});$size=$(if($r.byteLength){'{0:N1} MB' -f ($r.byteLength/1MB)}else{'-'});$hash=$(if($r.sha256){([string]$r.sha256).Substring(0,8)}else{'-'});$index=$(switch([string]$r.indexState){'Ready'{'Ready ('+$r.indexCount+' providers)'}'Stale'{'Needs rebuild'}'Unknown'{'Not built'}default{'-'}});$rows+='<tr><td>'+(ConvertTo-HtmlSafe $r.displayName)+'</td><td>'+$state+'</td><td>'+(ConvertTo-HtmlSafe $r.canonicalFileName)+'</td><td>'+$size+'</td><td>'+(ConvertTo-HtmlSafe $r.fileCreationTime)+'</td><td>'+(ConvertTo-HtmlSafe $r.lastWriteTime)+'</td><td>'+(ConvertTo-HtmlSafe $r.importedUtc)+'</td><td>'+(ConvertTo-HtmlSafe $r.rowCount)+'</td><td>'+$hash+'</td><td>'+(ConvertTo-HtmlSafe $index)+'</td><td><button onclick="pickFile()">Refresh</button></td></tr>'}
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>Provider Analysis</title>
<style>body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px}.card{background:white;border:1px solid #dce4ef;overflow:auto}__TABCSS__table{border-collapse:collapse;width:100%;min-width:1250px}th,td{padding:9px;border-bottom:1px solid #dce4ef;text-align:left;white-space:nowrap}th{background:#eaf1f8}button{padding:7px 11px;cursor:pointer}.stop{background:#a61b1b;color:white}.muted{color:#667085}.err{color:#a61b1b}#feeds table{min-width:0}#feeds td{white-space:normal}.modal{position:fixed;inset:0;background:#0008;display:none;align-items:center;justify-content:center}.box{background:white;padding:22px;border-radius:8px;max-width:650px;white-space:pre-wrap}</style></head>
<body><header><h1>Provider Analysis Automation</h1><p>Draft 5.4 - import the seven source exports by upload or report server feed; each is validated, fingerprinted and indexed</p></header>__NAV__
<main><button onclick="location.reload()">Reload</button> <button onclick="reindex()">Rebuild provider indexes</button> <button class="stop" onclick="stopServer()">Stop server</button><input id="file" type="file" accept=".xlsx" hidden>
<h2>Canonical sources</h2><div class="card"><table><thead><tr><th>Source</th><th>Status</th><th>File</th><th>Size</th><th>Export/File created</th><th>Last write</th><th>Imported UTC</th><th>Rows</th><th>SHA-256</th><th>Provider index</th><th>Action</th></tr></thead><tbody>__ROWS__</tbody></table></div>
<p>Uploads are staged and fingerprinted before replacement. Nothing changes until you confirm. The provider index for a source is rebuilt automatically after each import, so the wizard opens its name lists instantly.</p>
<h2>Report server feeds</h2><div class="card" style="padding:14px"><p class="muted">On a source's dashboard choose Export, <b>Data feed</b>, and save the .atomsvc file; add it here. The script downloads the same Excel export the dashboard produces, signed in as you, recognizes the source from its columns, and imports it through the usual validate, archive and index steps. Sync on demand, or daily at a set hour while the server is running.</p><div id="feeds">Loading...</div><p><button onclick="pickFeed()">Add feed (.atomsvc)...</button> <button onclick="syncNow('')">Sync all now</button> <label style="margin-left:14px"><input type="checkbox" id="schedEnabled" onchange="saveSchedule()"> Sync daily after</label> <select id="schedHour" onchange="saveSchedule()"></select> <span id="feedNote" class="muted"></span></p><input id="feedFile" type="file" accept=".atomsvc,.xml,.txt" hidden></div></main>
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
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function when(iso){if(!iso)return '';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleString()}
function hourLabel(h){return (h%12||12)+':00 '+(h<12?'AM':'PM')}
function feedsHtml(m){if(!m.feeds.length)return '<p class="muted">No feeds yet.</p>';return '<table><thead><tr><th>Source</th><th>Report</th><th>Parameters</th><th>Current file imported</th><th>Last sync</th><th></th></tr></thead><tbody>'+m.feeds.map(f=>'<tr><td><b>'+esc(f.displayName)+'</b></td><td title="'+esc(f.server)+'">'+esc(f.reportPath)+'</td><td class="muted">'+esc(f.parameterText)+'</td><td>'+(f.currentImportedUtc?esc(when(f.currentImportedUtc)):'<span class="muted">none</span>')+'</td><td>'+(f.lastSyncUtc?esc(when(f.lastSyncUtc))+'<br><span class="'+(f.lastResult.indexOf('Failed')===0?'err':'muted')+'">'+esc(f.lastResult)+'</span>':'<span class="muted">never</span>')+'</td><td><button onclick="syncNow(\''+esc(f.sourceKey)+'\')">Sync</button> <button onclick="removeFeed(\''+esc(f.sourceKey)+'\')">Remove</button></td></tr>').join('')+'</tbody></table>'}
async function loadFeeds(){try{const m=await call('/api/sync');document.getElementById('feeds').innerHTML=feedsHtml(m);const sel=document.getElementById('schedHour');if(!sel.options.length){for(let h=0;h<24;h++)sel.add(new Option(hourLabel(h),h))}sel.value=m.schedule.hour;document.getElementById('schedEnabled').checked=m.schedule.enabled;document.getElementById('feedNote').textContent=(m.schedule.enabled?'Daily sync runs after '+hourLabel(m.schedule.hour)+' while the server is running'+(m.schedule.lastRunDate?' (last ran '+m.schedule.lastRunDate+')':'')+'.':'')+(m.reportRunning?' A report is running; sync waits until it finishes.':'')}catch(e){document.getElementById('feeds').innerHTML='<p class="err">'+esc(e.message)+'</p>'}}
function pickFeed(){const f=document.getElementById('feedFile');f.value='';f.click()}
document.getElementById('feedFile').onchange=()=>{const f=document.getElementById('feedFile');if(!f.files.length)return;const file=f.files[0];const rd=new FileReader();rd.onload=async()=>{show('Reading '+file.name+' and downloading the report export to recognize the source (large reports can take a minute) ...',false);try{const j=await call('/api/sync-feed',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({text:rd.result})});token=j.staged.token;show('Feed registered for '+j.feed.displayName+'.\nReport: '+j.feed.reportPath+'\nParameters: '+j.feed.parameterText+'\nDownloaded '+j.staged.rowCount+' rows, SHA-256 '+j.staged.sha256+'.\n\nConfirm to import this download now, or Close and use Sync later.',true);loadFeeds()}catch(e){show('Error: '+e.message,false)}};rd.readAsText(file)};
async function syncNow(key){show('Syncing '+(key||'all feeds')+' from the report server (each source downloads, validates and, if changed, replaces the current file and rebuilds its index) ...',false);try{const r=await call('/api/sync-now'+(key?'?sourceKey='+encodeURIComponent(key):''),{method:'POST'});show(r.map(x=>x.displayName+': '+x.result+' - '+x.detail+' ('+x.seconds+'s)').join('\n'),false);loadFeeds();setTimeout(()=>{if(r.some(x=>x.result==='Replaced'))location.reload()},4000)}catch(e){show('Error: '+e.message,false)}}
async function removeFeed(key){if(!confirm('Remove the feed for '+key+'? The current file stays; only the download link is forgotten.'))return;try{await call('/api/sync-feed-remove?sourceKey='+encodeURIComponent(key),{method:'POST'});loadFeeds()}catch(e){show('Error: '+e.message,false)}}
async function saveSchedule(){try{await call('/api/sync-schedule',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({enabled:document.getElementById('schedEnabled').checked,hour:+document.getElementById('schedHour').value})});loadFeeds()}catch(e){show('Error: '+e.message,false)}}
loadFeeds();
</script></body></html>
'@
 return $html.Replace('__ROWS__',$rows).Replace('__NAV__',(Get-NavHtml 'sources')).Replace('__TABCSS__',$script:TabCss)
}
function Send($Context,[int]$Code,[string]$Type,[string]$Body){$b=[Text.Encoding]::UTF8.GetBytes($Body);$Context.Response.StatusCode=$Code;$Context.Response.ContentType=$Type;$Context.Response.ContentLength64=$b.Length;$Context.Response.Headers['Cache-Control']='no-store';$Context.Response.OutputStream.Write($b,0,$b.Length);$Context.Response.Close()}
function Request($Context){
 $remote=$null;try{$remote=$Context.Request.RemoteEndPoint}catch{}
 if($null -eq $remote -or -not ([Net.IPAddress]::IsLoopback($remote.Address))){try{Send $Context 403 'text/plain' 'Forbidden'}catch{};return}
 $method=$Context.Request.HttpMethod;$path=$Context.Request.Url.AbsolutePath.TrimEnd('/');if(!$path){$path='/'}
 try{if($method -eq 'GET' -and $path -eq '/'){Send $Context 200 'text/html; charset=utf-8' (Page);return};if($method -eq 'GET' -and $path -eq '/providers'){Send $Context 200 'text/html; charset=utf-8' (ProviderPage);return};if($method -eq 'GET' -and $path -eq '/provider-index'){Send $Context 200 'text/html; charset=utf-8' (ProviderIndexPage);return};if($method -eq 'GET' -and $path -eq '/api/locations'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-KnownLocations ([string]$Context.Request.QueryString['riskPool'])) -Depth 3);return};if($method -eq 'GET' -and $path -eq '/api/provider-index'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-ProviderIndexModel) -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/profile-job'){$result=New-JobFromProfile ([string]$Context.Request.QueryString['npi']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/health'){Send $Context 200 'application/json' '{"status":"ok","draft":4}';return};if($method -eq 'GET' -and $path -eq '/api/status'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Status) -Depth 6);return};if($method -eq 'GET' -and $path -eq '/api/provider-sources'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @((Json $script:ConfigPath).sources|Select-Object sourceKey,displayName,providerColumns) -Depth 5);return};if($method -eq 'GET' -and $path -eq '/api/providers'){$sourceKey=[string]$Context.Request.QueryString['sourceKey'];$sw=[Diagnostics.Stopwatch]::StartNew();$result=@(Get-ProviderIndexForPool $sourceKey ([string]$Context.Request.QueryString['riskPool']));Log 'PROVIDER_LIST' 'OK' ($result.Count.ToString()+' names in '+[Math]::Round($sw.Elapsed.TotalSeconds,2)+'s') $sourceKey;Send $Context 200 'application/json' (ConvertTo-Json -InputObject $result -Depth 6);return};if($method -eq 'GET' -and $path -eq '/api/suggest'){$names=@($Context.Request.QueryString.GetValues('name')|Where-Object{$_});$result=@(Get-ProviderSuggestions ([string]$Context.Request.QueryString['sourceKey']) $names ([string]$Context.Request.QueryString['riskPool']));Send $Context 200 'application/json' (ConvertTo-Json -InputObject $result -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/reindex'){$out=@();foreach($s in (Json $script:ConfigPath).sources){if(!(Test-Path -LiteralPath (Join-Path $script:Paths.CanonicalCurrent $s.canonicalFileName))){continue};$r=Update-ProviderIndexAfterImport ([string]$s.sourceKey) -Force;$out+=[ordered]@{sourceKey=$s.sourceKey;displayName=$s.displayName;count=$r.count;seconds=$r.seconds;error=$r.error}};Send $Context 200 'application/json' (ConvertTo-Json -InputObject @($out) -Depth 5);return};if($method -eq 'GET' -and $path -eq '/api/profiles'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-ProviderProfiles) -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/profile'){$result=Save-ProviderProfile (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/html'){Send-JobOutput $Context ([string]$Context.Request.QueryString['jobId']) 'html';return};if($method -eq 'GET' -and $path -eq '/pdf'){Send-JobOutput $Context ([string]$Context.Request.QueryString['jobId']) 'pdf';return};if($method -eq 'GET' -and $path -eq '/flag'){Send $Context 200 'text/html; charset=utf-8' (ConvertTo-AnalysisHtml (Get-FlagReportModel ([string]$Context.Request.QueryString['jobId'])) -Interactive);return};if($method -eq 'GET' -and $path -eq '/communication'){Send $Context 200 'text/html; charset=utf-8' (CommunicationPage);return};if($method -eq 'GET' -and $path -eq '/api/communication'){Send $Context 200 'application/json' ((Get-CommunicationModel)|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/contact-list'){$result=Select-ContactList ([string]$Context.Request.QueryString['path']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/contact-list-upload'){$result=Receive-ContactListUpload $Context;Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'GET' -and $path -eq '/api/recipients'){Send $Context 200 'application/json' ((Get-CommunicationRecipients ([string]$Context.Request.QueryString['riskPool']))|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign'){$result=New-Campaign (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/campaign'){Send $Context 200 'application/json' ((Get-CampaignModel ([string]$Context.Request.QueryString['id']))|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/campaigns'){Send $Context 200 'application/json' ([ordered]@{queue=(Get-QueueStatus);campaigns=@(Get-CampaignList)}|ConvertTo-Json -Depth 9);return};if($method -eq 'GET' -and $path -eq '/api/queue'){Send $Context 200 'application/json' ((Get-QueueStatus)|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/job-reset'){$result=Reset-RunningJob ([string]$Context.Request.QueryString['jobId']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign-cancel'){$result=Stop-Campaign ([string]$Context.Request.QueryString['id']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign-drafts'){$result=Invoke-CampaignDrafts ([string]$Context.Request.QueryString['id']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/campaign-resume'){$result=Resume-Campaign ([string]$Context.Request.QueryString['id']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/overview'){Send $Context 200 'text/html; charset=utf-8' (OverviewPage);return};if($method -eq 'GET' -and $path -eq '/api/overview'){Send $Context 200 'application/json' ((Get-OverviewPageModel)|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/overview'){$result=New-OverviewRun (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/overview-run'){Send $Context 200 'application/json' ((Get-OverviewModel ([string]$Context.Request.QueryString['id']))|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/overview.xlsx'){Send-OverviewFile $Context ([string]$Context.Request.QueryString['id']) 'xlsx';return};if($method -eq 'GET' -and $path -eq '/overview.pdf'){Send-OverviewFile $Context ([string]$Context.Request.QueryString['id']) 'pdf';return}
  if($method -eq 'GET' -and $path -eq '/new-patients'){Send $Context 200 'text/html; charset=utf-8' (NewPatientsPage);return};if($method -eq 'GET' -and $path -eq '/api/new-patients'){Send $Context 200 'application/json' ((Get-NewListPageModel)|ConvertTo-Json -Depth 8);return}
  if($method -eq 'POST' -and $path -eq '/api/new-patients'){$result=New-NewListRun (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/new-patients-run'){Send $Context 200 'application/json' ((Get-NewListModel ([string]$Context.Request.QueryString['id']))|ConvertTo-Json -Depth 8);return}
  if($method -eq 'GET' -and $path -eq '/new-patients.file'){Send-NewListFile $Context ([string]$Context.Request.QueryString['id']) ([string]$Context.Request.QueryString['n']);return};
  if($path -eq '/combine' -or $path.StartsWith('/combine/') -or $path.StartsWith('/api/combine/')){if(Invoke-CmbRequest $Context $method $path){return}}
  if($method -eq 'GET' -and $path -eq '/poc-coding'){Send $Context 200 'text/html; charset=utf-8' (PocCodingPage);return};if($method -eq 'GET' -and $path -eq '/api/poc-coding'){Send $Context 200 'application/json' ((Get-PocPageModel)|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/poc-coding-upload'){$result=Receive-PocUpload $Context;Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/poc-coding-file'){Send $Context 200 'application/json' ((Get-PocFileModel ([string]$Context.Request.QueryString['token']))|ConvertTo-Json -Depth 8);return}
  if($method -eq 'POST' -and $path -eq '/api/poc-coding'){$result=New-PocRun (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/poc-coding-run'){Send $Context 200 'application/json' ((Get-PocRunModel ([string]$Context.Request.QueryString['id']))|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/poc-coding.report'){Send-PocReport $Context ([string]$Context.Request.QueryString['id']) ([string]$Context.Request.QueryString['download'] -eq '1');return}
 if($method -eq 'GET' -and $path -eq '/api/sync'){Send $Context 200 'application/json' ((Get-SyncModel)|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/sync-feed'){$b=Read-BodyJson $Context;$result=Register-ReportFeed ([string](Get-P $b 'text' ''));Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/sync-feed-remove'){$result=Remove-ReportFeed ([string]$Context.Request.QueryString['sourceKey']);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/sync-now'){$result=Invoke-FeedSync @([string]$Context.Request.QueryString['sourceKey']);Send $Context 200 'application/json' (ConvertTo-Json -InputObject @($result) -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/sync-schedule'){$result=Set-SyncSchedule (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'GET' -and $path -eq '/api/flags'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-ProviderFlags ([string]$Context.Request.QueryString['npi'])) -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/flag'){$result=Save-FlagFromRequest (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/run-next'){$script:RunNext=$true;Send $Context 202 'application/json' '{"status":"scheduled"}';return};if($method -eq 'GET' -and $path -eq '/api/jobs'){Send $Context 200 'application/json' (ConvertTo-Json -InputObject @(Get-Jobs) -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/job'){$result=New-PreparedJob (Read-BodyJson $Context);Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 8);return};if($method -eq 'POST' -and $path -eq '/api/stage'){$result=Receive-Upload $Context;Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/confirm'){$result=Publish-StagedFile ([string]$Context.Request.QueryString['token']);$index=Update-ProviderIndexAfterImport ([string]$result.sourceKey) -Force;$result['providerCount']=$index.count;$result['indexSeconds']=$index.seconds;$result['indexError']=$index.error;Send $Context 200 'application/json' ($result|ConvertTo-Json -Depth 6);return};if($method -eq 'POST' -and $path -eq '/api/stop'){$script:Stop=$true;Send $Context 202 'application/json' '{"status":"stopping"}';return};Send $Context 404 'application/json' '{"error":"not found"}'}catch{$msg=Safe $_.Exception.Message;Log 'HTTP_REQUEST' 'FAILED' $msg;try{Send $Context 400 'application/json' (([ordered]@{error=$msg}|ConvertTo-Json -Compress))}catch{}}
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
 $tabs=@(@('sources','/','Sources'),@('wizard','/providers','Provider Wizard'),@('index','/provider-index','Provider Index'),@('communication','/communication','Communication'),@('overview','/overview','Overview'),@('newpatients','/new-patients','New Patients'),@('poc','/poc-coding','POC Coding'),@('combine','/combine','Combine'))
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
$script:Outlook=$null;$script:OutlookKeep=$null
function Test-OutlookAlive($Outlook){
 # A method call through to the mail store fails once Outlook has exited; reading a property of the old connection can quietly return nothing instead.
 try{$ns=$Outlook.GetNamespace('MAPI');if($null -eq $ns){return $false};return ($null -ne $ns.GetDefaultFolder(16))}catch{return $false}
}
function Reset-OutlookApplication{
 foreach($o in @($script:OutlookKeep,$script:Outlook)){if($null -ne $o){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}}}
 $script:OutlookKeep=$null;$script:Outlook=$null
}
function Get-OutlookApplication{
 # The connection is kept between drafts but checked each time: Outlook started in the background (no window open) can shut itself down after
 # a draft is closed, and a stale connection then fails every later draft ("The RPC server is unavailable"). A dead one is replaced, waiting a
 # moment for a closing Outlook to finish exiting. A hidden Drafts window object is held while connected, which keeps a background Outlook running.
 if($script:Outlook -and (Test-OutlookAlive $script:Outlook)){return $script:Outlook}
 Reset-OutlookApplication;$err=''
 for($try=1;$try -le 3;$try++){
  try{$script:Outlook=[Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')}catch{$script:Outlook=$null}
  if(!$script:Outlook){try{$script:Outlook=New-Object -ComObject Outlook.Application}catch{$script:Outlook=$null;$err=$_.Exception.Message}}
  if($script:Outlook -and (Test-OutlookAlive $script:Outlook)){
   try{$script:OutlookKeep=$script:Outlook.GetNamespace('MAPI').GetDefaultFolder(16).GetExplorer()}catch{$script:OutlookKeep=$null}
   return $script:Outlook
  }
  Reset-OutlookApplication;Start-Sleep -Seconds (2*$try)
 }
 throw ('Outlook could not be started for drafting'+$(if($err){': '+$err}else{'; open Outlook and try again.'}))
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
 # A lost Outlook connection (Outlook closed or restarted between drafts) is reconnected and the draft tried again, up to three times.
 for($attempt=1;;$attempt++){
  try{return (New-OutlookDraftItem $Spec)}
  catch{
   $m=$_.Exception.Message;$lost=($m -match '0x800706BA|0x800706BE|0x80010108|0x80010114|0x800401FD|RPC server|disconnected|null-valued|did not create')
   if(!$lost){throw};Reset-OutlookApplication
   if($attempt -ge 3){throw ('Outlook closed or stopped responding while drafting ('+$m+'). Open Outlook, then select Retry failed items.')}
   Log 'OUTLOOK_RECONNECT' 'WARN' ('Attempt '+$attempt+': '+$m);Start-Sleep -Seconds (2*$attempt)
  }
 }
}
function New-OutlookDraftItem($Spec){
 $ol=Get-OutlookApplication;$mail=$ol.CreateItem(0);$tempDir=''
 if($null -eq $mail){throw 'Outlook did not create a mail item.'}
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
function Get-SourceRows([string]$Key,[string[]]$Wanted,[string]$FilterHeader,[string]$FilterValue,[hashtable]$FilterSet=$null){
 # FilterSet (keys = provider names) keeps rows for several providers in one pass; the caller groups them by the filter column.
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
   $filtering=[bool]($FilterHeader -and ($FilterValue -or $null -ne $FilterSet));$target=$(if($filtering){$FilterValue.Trim()}else{''})
   $filterVals=$(if($filtering -and $bulk.ContainsKey($FilterHeader)){$bulk[$FilterHeader]}else{$null})
   for($i=0;$i -lt $count;$i++){
    if($filtering){$fv=$(if($null -ne $filterVals){$filterVals[$i]}else{''});if($null -ne $FilterSet){if(!$FilterSet.ContainsKey([string]$fv)){continue}}elseif($fv -ne $target){continue}}
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
# Fields each source contributes to patient facts, and the provider column its rows are filtered on (one report or a whole overview reads the same columns).
$script:FactFields=[ordered]@{
 Export=@('Provider Name','Risk Pool','NPI','MemberID','First Name','Last Name','Date of Birth','Payer','Last QEM date with non-PCP','Last ACV Date','Last QEM Visit Date with any PCP in assigned TIN','Completed Attestations','Incompleted Attestations','Open ICDs','New Patient','Total Care Gaps','TCM','Phone Number','Active Date')
 PtListQuality=@('Provider Name','Risk Pool','NPI','MemberID','First Name','Last Name','DOB','BCS','COLO','EED','GSD','CBP','OMW','KED','SPC','MAD','MAC','MAH','SUPD','COB','POLY','OMW Critical Due Date')
 DiabetesScorecard=@('Provider','Cdo','Patient','Member ID','Mrn','KED','EED','Eye Exam Gap Status','Eye Exam Date','Next Appt Date','Next Appt Specialty','Next Appt Location','Risk','GSD','Med Adherence DM','MAD Days Supply','Dx Date','Avg Last A1c','% eGFR last 12 mo.','% uACR last 12 mo.')
 SerialScheduling=@('Provider Name','Risk Pool','Member ID','MRN','Future PCP Visits 2026','PCP Visit Dates','A1c Date')
 HR=@('PCP Name','Patient First Name','Patient Last Name','DOB','Patient Insurance ID')
 RiskPopulationOutreach=@('Epic Pcp','Cdo','Member ID','Mrn','Next Acv')
}
$script:FactFilters=[ordered]@{Export='Provider Name';PtListQuality='Provider Name';DiabetesScorecard='Provider';SerialScheduling='Provider Name';RiskPopulationOutreach='Epic Pcp';HR='PCP Name'}
function Get-ProviderPatientFacts($Job){
 # Reads the seven sources for one provider and returns plain patient facts (no scoring, no tables) so a report can be rebuilt later without Excel.
 $a=$Job.aliases;$exportAlias=Get-Alias $a 'Export';$qualityAlias=Get-Alias $a 'PtListQuality';$serialAlias=Get-Alias $a 'SerialScheduling';$dscAlias=Get-Alias $a 'DiabetesScorecard';$rpoAlias=Get-Alias $a 'RiskPopulationOutreach';$hrKey=$(if(Get-Alias $a 'HR-CRH'){'HR-CRH'}elseif(Get-Alias $a 'HR-RIVPHNYCMM'){'HR-RIVPHNYCMM'}else{''});$hrAlias=$(if($hrKey){Get-Alias $a $hrKey}else{''})
 $exports=@(Get-SourceRows 'Export' $script:FactFields.Export 'Provider Name' $exportAlias);Set-JobProgress $Job 15 'Loaded Export source';$quality=@(Get-SourceRows 'PtListQuality' $script:FactFields.PtListQuality 'Provider Name' $qualityAlias);Set-JobProgress $Job 28 'Loaded PtListQuality source';$dsc=@();if($dscAlias){$dsc=@(Get-SourceRows 'DiabetesScorecard' $script:FactFields.DiabetesScorecard 'Provider' $dscAlias)};Set-JobProgress $Job 40 'Loaded Diabetes Scorecard source';$serial=@();if($serialAlias){$serial=@(Get-SourceRows 'SerialScheduling' $script:FactFields.SerialScheduling 'Provider Name' $serialAlias)};Set-JobProgress $Job 50 'Loaded Serial Scheduling source';$rpo=@();if($rpoAlias){$rpo=@(Get-SourceRows 'RiskPopulationOutreach' $script:FactFields.RiskPopulationOutreach 'Epic Pcp' $rpoAlias)};Set-JobProgress $Job 54 'Loaded Risk Population Outreach source';$hr=@();if($hrKey -and $hrAlias){$hr=@(Get-SourceRows $hrKey $script:FactFields.HR 'PCP Name' $hrAlias)};Set-JobProgress $Job 58 'Loaded HR source'
 return (New-PatientFacts $exports $quality $dsc $serial $rpo $hr)
}
function New-PatientFacts([object[]]$exports,[object[]]$quality,[object[]]$dsc,[object[]]$serial,[object[]]$rpo,[object[]]$hr){
 # Joins one provider's rows from the seven sources into plain patient facts; used by a single report and by the group overview.
 $exports=@($exports|Where-Object{$null -ne $_});$quality=@($quality|Where-Object{$null -ne $_});$dsc=@($dsc|Where-Object{$null -ne $_});$serial=@($serial|Where-Object{$null -ne $_});$rpo=@($rpo|Where-Object{$null -ne $_});$hr=@($hr|Where-Object{$null -ne $_})
 $qIndex=New-UniqueIndex $quality 'MemberID';$dIndex=New-UniqueIndex $dsc 'Member ID';$sIndex=New-UniqueIndex $serial 'Member ID';$rpoIndex=New-UniqueIndex $rpo 'Member ID';$hrIds=@{};foreach($x in $hr){$id=([string]$x.'Patient Insurance ID').Trim();if($id){$hrIds[$id]=$x}};$exportIds=New-Object Collections.Generic.List[string];foreach($x in $exports){$exportIds.Add(([string]$x.MemberID).Trim())};$highIds=@{};$unmatchedHr=@{};foreach($hid in $hrIds.Keys){$matchCount=0;$matchId='';foreach($eid in $exportIds){if($eid -eq $hid -or $eid.StartsWith($hid)){$matchCount++;if($matchCount -eq 1){$matchId=$eid}}};if($matchCount -eq 1){$highIds[$matchId]=$true}elseif($matchCount -eq 0){$unmatchedHr[$hid]=$hrIds[$hid]}}
 $hedis=@('BCS','COLO','EED','GSD','CBP','OMW','KED','SPC','MAD','MAC','MAH','SUPD','COB','POLY');$patients=@();$seen=@{}
 foreach($e in $exports){$id=([string]$e.MemberID).Trim();if(!$id){continue};$seen[$id]=$true;$q=$(if($qIndex.ContainsKey($id)){$qIndex[$id]}else{$null});$d=$(if($dIndex.ContainsKey($id)){$dIndex[$id]}else{$null});$s=$(if($sIndex.ContainsKey($id)){$sIndex[$id]}else{$null});$rp=$(if($rpoIndex.ContainsKey($id)){$rpoIndex[$id]}else{$null})
  $open=@();$closed=@();if($q){foreach($m in $hedis){$v=([string]$q.$m).Trim();if($v -in @('Needed','Non-compliant')){$open+=$m}elseif($v -in @('Completed','Compliant')){$closed+=$m}}}
  if($d){$overrides=@{'KED'=[string]($d.KED);'EED'=[string]($d.EED);'GSD'=[string]($d.GSD);'MAD'=[string]($d.'Med Adherence DM')};foreach($m in $overrides.Keys){$v=$overrides[$m];if($v -in @('Needed','Non-compliant')){if($open -notcontains $m){$open+=$m};$closed=@($closed|Where-Object{$_ -ne $m})}elseif($v -in @('Completed','Compliant')){$open=@($open|Where-Object{$_ -ne $m});if($closed -notcontains $m){$closed+=$m}}};$egfrComplete=(ConvertTo-NumberValue $d.'% eGFR last 12 mo.') -ge 100;$uacrComplete=(ConvertTo-NumberValue $d.'% uACR last 12 mo.') -ge 100;if($egfrComplete -and $uacrComplete){$open=@($open|Where-Object{$_ -ne 'KED'});if($closed -notcontains 'KED'){$closed+='KED'}}}
  $high=$highIds.ContainsKey($id);$lastAcv=(ConvertTo-DateValue $e.'Last ACV Date');$acor=([string]$e.Payer).Trim() -eq 'ACOR';$due=$false;if($lastAcv){$due=$(if($acor){$lastAcv -le (Get-Date).AddDays(-366)}else{$lastAcv -lt (Get-Date -Day 1 -Month 1)})};$nextAcv=$null;if($rp){$nextAcv=ConvertTo-DateValue $rp.'Next Acv'};if($nextAcv -and $nextAcv.Date -ge (Get-Date).Date){$due=$false}
  $vis=@((ConvertTo-DateValue $e.'Last QEM date with non-PCP'),(ConvertTo-DateValue $e.'Last ACV Date'),(ConvertTo-DateValue $e.'Last QEM Visit Date with any PCP in assigned TIN'))|Where-Object{$_};$last=$(if($vis){$vis|Sort-Object -Descending|Select-Object -First 1}else{$null});$serialLast=$(if($s){Get-LatestDateFromText ([string]$s.'PCP Visit Dates')}else{$null})
  $mrn='';foreach($cand in @($(if($d){$d.Mrn}else{''}),$(if($s){$s.MRN}else{''}),$(if($rp){$rp.Mrn}else{''}))){$c=([string]$cand).Trim();if($c -and $c -ne '0'){$mrn=$c;break}}
  $patients+=[pscustomobject][ordered]@{MemberID=$id;MRN=$mrn;First=[string]$e.'First Name';Last=[string]$e.'Last Name';DOB=$(ConvertTo-DateText $e.'Date of Birth');ACOR=$acor;ACVDue=$due;Attest=((ConvertTo-NumberValue $e.'Completed Attestations')+(ConvertTo-NumberValue $e.'Incompleted Attestations')) -gt 0;OpenICD=[string]$e.'Open ICDs';NewPatient=(Test-TrueValue $e.'New Patient');OpenHedis=($open -join ', ');ClosedHedis=($closed -join ', ');HighRisk=$high;LastVisit=$last;SerialLast=$serialLast;A1cDate=$(if($s){ConvertTo-DateValue $s.'A1c Date'}else{$null});FutureVisits=$(if($s){[string]$s.'Future PCP Visits 2026'}else{''});NextAppt=$(if($d){ConvertTo-DateValue $d.'Next Appt Date'}else{$null});NextApptSpecialty=$(if($d){([string]$d.'Next Appt Specialty').Trim()}else{''});A1c=$(if($d){ConvertTo-NumberOrNull $d.'Avg Last A1c'}else{$null});A1cText=$(if($d){([string]$d.'Avg Last A1c').Trim()}else{''});Diabetic=($null -ne $d);EyeOpen=[bool]($d -and (([string]$d.'Eye Exam Gap Status').Trim() -eq 'Open'));EyeDate=$(if($d){ConvertTo-DateValue $d.'Eye Exam Date'}else{$null});EgfrNeeded=[bool]($d -and ((ConvertTo-NumberValue $d.'% eGFR last 12 mo.') -lt 100));UacrNeeded=[bool]($d -and ((ConvertTo-NumberValue $d.'% uACR last 12 mo.') -lt 100));MedAdhDM=$(if($d){([string]$d.'Med Adherence DM').Trim()}else{''});OmwDue=$(if($q){ConvertTo-DateText $q.'OMW Critical Due Date'}else{''});TotalGaps=(ConvertTo-NumberValue $e.'Total Care Gaps');IncompleteAttest=(ConvertTo-NumberValue $e.'Incompleted Attestations');OpenIcdCount=@(([string]$e.'Open ICDs') -split ','|Where-Object{$_.Trim() -and $_.Trim() -ne '0'}).Count;Tcm=[bool]((Test-TrueValue $e.TCM) -or ((ConvertTo-NumberValue $e.TCM) -gt 0));RiskContract=$true;DSC=$d;Payer=([string]$e.Payer).Trim();Phone=([string](Get-Field $e 'Phone Number' '')).Trim();ActiveDate=(ConvertTo-DateText (Get-Field $e 'Active Date' ''));PcpVisit=(ConvertTo-DateValue $e.'Last QEM Visit Date with any PCP in assigned TIN');NextAcv=$nextAcv}
 }
 foreach($hid in $unmatchedHr.Keys){$x=$unmatchedHr[$hid];$patients+=[pscustomobject][ordered]@{MemberID=$hid;MRN='';First=[string]$x.'Patient First Name';Last=[string]$x.'Patient Last Name';DOB=$(ConvertTo-DateText $x.DOB);ACOR=$null;ACVDue=$false;Attest=$false;OpenICD='';NewPatient=$false;OpenHedis='';ClosedHedis='';HighRisk=$true;LastVisit=$null;SerialLast=$null;A1cDate=$null;FutureVisits='';NextAppt=$null;NextApptSpecialty='';A1c=$null;A1cText='';Diabetic=$false;EyeOpen=$false;EyeDate=$null;EgfrNeeded=$false;UacrNeeded=$false;MedAdhDM='';OmwDue='';TotalGaps=0;IncompleteAttest=0;OpenIcdCount=0;Tcm=$false;RiskContract=$true;DSC=$null;Payer='';Phone='';ActiveDate='';PcpVisit=$null;NextAcv=$null}}
 return [ordered]@{asOf=(Get-Date).ToString('o');riskPool=$(if($exports.Count){[string]$exports[0].'Risk Pool'}else{''});patients=@($patients);scorecardNonRisk=@($dsc|Where-Object{([string]$_.Risk) -match 'Non-Risk Population'})}
}
$script:PatientDateFields=@('LastVisit','SerialLast','A1cDate','NextAppt','EyeDate','PcpVisit','NextAcv')
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
# The dashboard's top-section counts, shared by each report and the group overview (flagged patients are already removed from $Active).
$script:KpiColumns=@(@('total','Total Patients'),@('acor','ACOR Payor'),@('nonAcor','Non-ACOR Payor'),@('acorAcvDue','ACOR + ACV Due'),@('nonAcorAcvDue','Non-ACOR + ACV Due'),@('newPatients','New Patients'),@('hasAttestations','Has Attestations'),@('openIcd','Open ICDs Present'),@('hasHedis','Has HEDIS'),@('openHedis','Open HEDIS Present'),@('highRisk','High Risk'),@('flagged','Flagged'))
function New-KpiSummary([object[]]$Active,[int]$Flagged){
 $active=@($Active|Where-Object{$null -ne $_});$total=$active.Count;$acorCount=@($active|Where-Object{$_.ACOR -eq $true}).Count;$non=@($active|Where-Object{$null -ne $_.ACOR -and $_.ACOR -eq $false}).Count
 return [ordered]@{total=$total;acor=$acorCount;nonAcor=$non;acorAcvDue=@($active|Where-Object{$_.ACOR -eq $true -and $_.ACVDue}).Count;nonAcorAcvDue=@($active|Where-Object{$_.ACOR -eq $false -and $_.ACVDue}).Count;newPatients=@($active|Where-Object{$_.NewPatient}).Count;hasAttestations=@($active|Where-Object{$_.Attest}).Count;openIcd=@($active|Where-Object{$_.OpenICD}).Count;hasHedis=@($active|Where-Object{$_.OpenHedis -or $_.ClosedHedis}).Count;openHedis=@($active|Where-Object{$_.OpenHedis}).Count;highRisk=@($active|Where-Object{$_.HighRisk}).Count;flagged=$Flagged}
}
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
   $kpi=New-KpiSummary $active $applied.Count;$total=$kpi.total;$acorCount=$kpi.acor;$non=$kpi.nonAcor
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
 $k=$Model.kpi;$tiles=@();foreach($c in $script:KpiColumns){$tiles+=$c[1];$tiles+=(Get-Field $k $c[0] 0)};$tileHtml='';for($i=0;$i -lt $tiles.Count;$i+=2){$tileHtml+='<div class="tile"><span>'+ (ConvertTo-HtmlEncoded $tiles[$i]) +'</span><strong>'+ (ConvertTo-HtmlEncoded $tiles[$i+1]) +'</strong></div>'}
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
# --- Draft 5.4: report server feeds - register an SSRS .atomsvc per source, then pull its Excel export on demand or daily ---
$script:SyncFormat='EXCELOPENXML'
function Get-SyncStorePath{return (Join-Path $script:Paths.State 'sync.json')}
function Get-SyncStore{
 $s=Json (Get-SyncStorePath);$store=[ordered]@{version=1;feeds=[ordered]@{};schedule=[ordered]@{enabled=$false;hour=6;lastRunDate=''};updatedUtc=''}
 if($s){$f=Get-P $s 'feeds' $null;if($f){foreach($p in $f.PSObject.Properties){$store.feeds[$p.Name]=$p.Value}};$sch=Get-P $s 'schedule' $null;if($sch){$store.schedule.enabled=[bool](Get-P $sch 'enabled' $false);$store.schedule.hour=[int](Get-P $sch 'hour' 6);$store.schedule.lastRunDate=[string](Get-P $sch 'lastRunDate' '')}}
 return $store
}
function Save-SyncStore($Store){$Store.updatedUtc=[DateTime]::UtcNow.ToString('o');Save-JsonAtomic (Get-SyncStorePath) $Store}
function ConvertFrom-AtomServiceDocument([string]$Text){
 # Returns the first collection href of an SSRS .atomsvc service document (the feed URL carrying the report path and parameters). A bare URL is accepted too.
 if([string]::IsNullOrWhiteSpace($Text)){throw 'The .atomsvc file is empty.'}
 $t=$Text.Trim();if($t -match '^https?://'){return $t}
 $xml=$null;try{$xml=[xml]$t}catch{throw 'That is not an .atomsvc service document (the XML could not be read).'}
 $hrefs=@($xml.SelectNodes('//*[local-name()="collection"]/@href')|ForEach-Object{[string]$_.Value}|Where-Object{$_})
 if($hrefs.Count -eq 0){throw 'No data feed collection was found in the .atomsvc file.'}
 return $hrefs[0]
}
function ConvertTo-ReportFeed([string]$Href){
 # Splits an SSRS render URL into server, report path and report parameters; rs:/rc: rendering switches are dropped so any export format can be requested later.
 $u=$null;try{$u=[Uri]$Href}catch{throw ('The feed URL is not valid: '+$Href)}
 if($u.Scheme -notin @('http','https')){throw 'The feed URL must start with http:// or https://.'}
 $query=$u.Query.TrimStart('?');if(!$query){throw 'The feed URL has no report path.'}
 $rawPath='';$reportPath='';$keep=@();$params=@()
 foreach($part in @($query -split '&'|Where-Object{$_ -ne ''})){
  $eq=$part.IndexOf('=');$rawKey=$(if($eq -ge 0){$part.Substring(0,$eq)}else{$part});$key=[Uri]::UnescapeDataString($rawKey.Replace('+',' '))
  if($eq -lt 0 -and !$reportPath -and $key.StartsWith('/')){$rawPath=$part;$reportPath=$key;continue}
  if($key -like 'rs:*' -or $key -like 'rc:*'){continue}
  $keep+=$part;$params+=[ordered]@{name=$key;value=$(if($eq -ge 0){[Uri]::UnescapeDataString($part.Substring($eq+1).Replace('+',' '))}else{''})}
 }
 if(!$reportPath){throw 'The feed URL has no report path (expected ...ReportServer?/Folder/Report&...).'}
 $base=$u.GetLeftPart([UriPartial]::Path)+'?'+$rawPath+$(if($keep.Count -gt 0){'&'+($keep -join '&')}else{''})
 return [ordered]@{server=$u.GetLeftPart([UriPartial]::Authority);reportPath=$reportPath;reportName=[string]$reportPath.Split('/')[-1];parameters=@($params);parameterText=(@($params|ForEach-Object{$_.name+'='+$_.value}) -join '; ');baseUrl=$base}
}
function Get-ReportRenderUrl($Feed,[string]$Format){return ([string](Get-Field $Feed 'baseUrl' '')+'&rs:Command=Render&rs:Format='+$Format)}
function Invoke-ReportDownload($Feed,[string]$Format,[string]$OutFile){
 # Fetches an export with the logged-in Windows user's credentials (no stored password) and checks that an Excel file, not an error or sign-in page, came back.
 $url=Get-ReportRenderUrl $Feed $Format;$reportPath=[string](Get-Field $Feed 'reportPath' '')
 # Windows PowerShell 5.1 sends the Windows sign-in over http without complaint; PowerShell 7 wants AllowUnencryptedAuthentication for that, so it is added only where the cmdlet has it.
 $args=@{Uri=$url;UseDefaultCredentials=$true;UseBasicParsing=$true;OutFile=$OutFile;TimeoutSec=900;ErrorAction='Stop'};if((Get-Command Invoke-WebRequest).Parameters.ContainsKey('AllowUnencryptedAuthentication')){$args.AllowUnencryptedAuthentication=$true}
 try{Invoke-WebRequest @args|Out-Null}
 catch{$code=0;try{$code=[int]$_.Exception.Response.StatusCode}catch{};if($code -eq 401 -or $code -eq 403){throw ('The report server refused the Windows sign-in ('+$code+'). Open the report in your browser first; if it asks for a password there, scripted sync is not possible with this server.')};if($code -eq 404){throw ('The report server did not find the report '+$reportPath+' (404).')};if($code -eq 500){throw ('The report server could not render '+$reportPath+' (500), usually a parameter problem; open the feed URL in your browser to see its message.')};throw ('Download failed: '+$_.Exception.Message)}
 if(!(Test-Path -LiteralPath $OutFile) -or (Get-Item -LiteralPath $OutFile).Length -lt 100){throw 'The report server returned an empty file.'}
 $head=New-Object byte[] 2;$fs=[IO.File]::OpenRead($OutFile);try{$null=$fs.Read($head,0,2)}finally{$fs.Dispose()}
 if($head[0] -ne 0x50 -or $head[1] -ne 0x4B){throw 'The report server returned a web page instead of an Excel file (probably a sign-in page or a report error). Open the feed URL in your browser to check.'}
}
function Register-ReportFeed([string]$Text){
 # Downloads the feed's Excel export once, lets the schema fingerprint decide which source it is, and stores the feed under that source. Returns the staged file too so it can be imported right away.
 $href=ConvertFrom-AtomServiceDocument $Text;$feed=ConvertTo-ReportFeed $href
 $token=[Guid]::NewGuid().ToString('N');$path=Join-Path $script:Paths.Staging ($token+'.xlsx')
 try{Invoke-ReportDownload $feed $script:SyncFormat $path}catch{Remove-Item $path -Force -ErrorAction SilentlyContinue;throw}
 $meta=Register-StagedFile $token $path ($feed.reportName+'.xlsx') 'report-server'
 $store=Get-SyncStore;$key=[string]$meta.sourceKey;$existing=$(if($store.feeds.Contains($key)){$store.feeds[$key]}else{$null})
 $record=[ordered]@{sourceKey=$key;displayName=[string]$meta.displayName;server=$feed.server;reportPath=$feed.reportPath;reportName=$feed.reportName;parameters=@($feed.parameters);parameterText=$feed.parameterText;baseUrl=$feed.baseUrl;addedUtc=$(if($existing){[string](Get-Field $existing 'addedUtc' '')}else{[DateTime]::UtcNow.ToString('o')});lastSyncUtc='';lastResult='';lastSha256='';lastRowCount=0}
 $store.feeds[$key]=$record;Save-SyncStore $store;Log 'FEED_REGISTERED' 'OK' ($key+' <- '+$feed.reportPath) $key
 return [ordered]@{feed=$record;staged=$meta}
}
function Remove-ReportFeed([string]$SourceKey){$store=Get-SyncStore;if(!$store.feeds.Contains($SourceKey)){throw 'No feed is registered for that source.'};$store.feeds.Remove($SourceKey);Save-SyncStore $store;Log 'FEED_REMOVED' 'OK' $SourceKey $SourceKey;return (Get-SyncModel)}
function Set-SyncSchedule($Body){$store=Get-SyncStore;$store.schedule.enabled=[bool](Get-P $Body 'enabled' $false);$h=[int](Get-P $Body 'hour' 6);if($h -lt 0 -or $h -gt 23){throw 'Hour must be 0-23.'};$store.schedule.hour=$h;Save-SyncStore $store;Log 'FEED_SCHEDULE' 'OK' ('enabled='+$store.schedule.enabled+' hour='+$h);return (Get-SyncModel)}
function Test-ReportRunning{return (@(Get-Jobs|Where-Object{[string](Get-P $_ 'state' '') -in @('Starting','Running')}).Count -gt 0)}
function Invoke-FeedSync([string[]]$SourceKeys){
 # Downloads each registered feed's Excel export; a file whose hash differs from the current canonical file replaces it through the normal staging, archive and index path.
 $store=Get-SyncStore;$keys=@($store.feeds.Keys);if($SourceKeys -and @($SourceKeys|Where-Object{$_}).Count -gt 0){$keys=@($keys|Where-Object{$SourceKeys -contains $_})}
 if($keys.Count -eq 0){throw 'No report feeds are registered.'}
 if(Test-ReportRunning){throw 'A report is being generated right now; sync again when it finishes so its source files are not replaced mid-run.'}
 $manifest=Json $script:ManifestPath;$results=@()
 foreach($key in $keys){
  $feed=$store.feeds[$key];$sw=[Diagnostics.Stopwatch]::StartNew();$r=[ordered]@{sourceKey=$key;displayName=[string](Get-Field $feed 'displayName' $key);result='';detail='';rowCount=0;seconds=0}
  $token=[Guid]::NewGuid().ToString('N');$path=Join-Path $script:Paths.Staging ($token+'.xlsx');$metaPath=Join-Path $script:Paths.Staging ($token+'.json')
  try{
   Invoke-ReportDownload $feed $script:SyncFormat $path
   $meta=Register-StagedFile $token $path ([string](Get-Field $feed 'reportName' $key)+'.xlsx') 'report-server'
   if([string]$meta.sourceKey -ne $key){Remove-Item $path,$metaPath -Force -ErrorAction SilentlyContinue;throw ('The download was recognized as '+$meta.sourceKey+', not '+$key+'; remove this feed and add the right .atomsvc.')}
   $current=$null;if($manifest){$src=Get-P $manifest 'sources' $null;if($src){$current=Get-P $src $key $null}}
   $r.rowCount=[int]$meta.rowCount
   if($current -and [string](Get-Field $current 'sha256' '') -eq [string]$meta.sha256){Remove-Item $path,$metaPath -Force -ErrorAction SilentlyContinue;$r.result='Unchanged';$r.detail='Same content as the current file.'}
   else{$pub=Publish-StagedFile $token;$r.result='Replaced';$r.detail=('Imported '+$meta.rowCount+' rows'+$(if((Get-Field $pub 'indexError' '')){'; index error: '+[string]$pub.indexError}else{'; provider index rebuilt'}))}
   Set-P $feed 'lastSha256' ([string]$meta.sha256);Set-P $feed 'lastRowCount' ([int]$meta.rowCount)
  }catch{Remove-Item $path,$metaPath -Force -ErrorAction SilentlyContinue;$r.result='Failed';$r.detail=$_.Exception.Message;Log 'FEED_SYNC' 'FAILED' ($key+': '+$_.Exception.Message) $key}
  $r.seconds=[Math]::Round($sw.Elapsed.TotalSeconds,1);Set-P $feed 'lastSyncUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $feed 'lastResult' ($r.result+': '+$r.detail);$results+=$r
 }
 Save-SyncStore $store;Log 'FEED_SYNC' 'OK' ((@($results|ForEach-Object{$_.sourceKey+'='+$_.result})) -join ', ')
 return $results
}
function Invoke-ScheduledFeedSync{
 # Heartbeat hook: once a day at or after the configured hour, when no report is running. The day is marked first so a failure is not retried every 15 seconds.
 $store=Get-SyncStore;if(!$store.schedule.enabled -or $store.feeds.Count -eq 0){return}
 $today=(Get-Date).ToString('yyyy-MM-dd');if($store.schedule.lastRunDate -eq $today -or (Get-Date).Hour -lt [int]$store.schedule.hour){return}
 if(Test-ReportRunning){return}
 $store.schedule.lastRunDate=$today;Save-SyncStore $store
 try{$null=Invoke-FeedSync @()}catch{Log 'FEED_SYNC' 'FAILED' $_.Exception.Message}
}
function Get-SyncModel{
 $store=Get-SyncStore;$manifest=Json $script:ManifestPath;$feeds=@()
 foreach($key in @($store.feeds.Keys)){$f=$store.feeds[$key];$current=$null;if($manifest){$src=Get-P $manifest 'sources' $null;if($src){$current=Get-P $src $key $null}};$feeds+=[ordered]@{sourceKey=$key;displayName=[string](Get-Field $f 'displayName' $key);server=[string](Get-Field $f 'server' '');reportPath=[string](Get-Field $f 'reportPath' '');parameterText=[string](Get-Field $f 'parameterText' '');addedUtc=(ConvertTo-IsoText (Get-Field $f 'addedUtc' ''));lastSyncUtc=(ConvertTo-IsoText (Get-Field $f 'lastSyncUtc' ''));lastResult=[string](Get-Field $f 'lastResult' '');currentImportedUtc=$(if($current){ConvertTo-IsoText (Get-Field $current 'importedUtc' '')}else{''})}}
 return [ordered]@{feeds=@($feeds|Sort-Object -Property @{Expression={[string]$_.displayName}});schedule=[ordered]@{enabled=[bool]$store.schedule.enabled;hour=[int]$store.schedule.hour;lastRunDate=[string]$store.schedule.lastRunDate};reportRunning=(Test-ReportRunning);user=[Environment]::UserName}
}
# --- Draft 5.5: Overview tab - the dashboard's top-section counts for a group of providers, one row each, with location, risk pool and grand totals ---
$script:OverviewKeep=30;$script:OverviewInline=$false
$script:OverviewSources=@('Export','PtListQuality','DiabetesScorecard','SerialScheduling','RiskPopulationOutreach','HR-CRH','HR-RIVPHNYCMM')
function Get-OverviewPath([string]$Id){if($Id -notmatch '^[a-f0-9]{32}$'){throw 'Invalid overview ID.'};return (Join-Path $script:Paths.State ('overview-'+$Id+'.json'))}
function Get-OverviewProviders{
 # Every saved profile, with the location and risk pool the Overview groups by.
 $out=@()
 foreach($p in @(Get-ProviderProfiles)){
  if($null -eq $p){continue};$npi=[string](Get-P $p 'npi' '');if($npi -notmatch '^\d{10}$'){continue}
  $loc=([string](Get-P $p 'location' '')).Trim();$pool=([string](Get-P $p 'riskPool' '')).Trim()
  $out+=[ordered]@{npi=$npi;displayName=[string](Get-P $p 'displayName' $npi);location=$(if($loc){$loc}else{'(No location)'});riskPool=$(if($pool){$pool}else{'(No risk pool)'})}
 }
 return @($out)
}
function Get-HrSourceKey($Aliases){if(Get-Alias $Aliases 'HR-CRH'){return 'HR-CRH'};if(Get-Alias $Aliases 'HR-RIVPHNYCMM'){return 'HR-RIVPHNYCMM'};return ''}
function Get-OverviewAlias($Aliases,[string]$Key){
 # The alias a provider's rows carry in one source; the HR sources follow the same one-of-two rule as a single report.
 if($Key -like 'HR-*'){if((Get-HrSourceKey $Aliases) -ne $Key){return ''}}
 return (Get-Alias $Aliases $Key).Trim()
}
function Get-GroupedRows($Groups,[string]$Key,[string]$Alias){if(!$Alias -or !$Groups.ContainsKey($Key)){return @()};$g=$Groups[$Key];if($g.ContainsKey($Alias)){return @($g[$Alias].ToArray())};return @()}
function Get-ActivePatients($Facts,[object[]]$Flags){
 # Flagged patients are resolved exactly as New-AnalysisModel does and left out; returns the rest and how many were flagged.
 $patients=@(Get-Field $Facts 'patients' @()|Where-Object{$null -ne $_});$applied=@{}
 if(@($Flags|Where-Object{$null -ne $_}).Count -gt 0){
  $prev=$script:AsOf;$script:AsOf=$(ConvertTo-DateValue (Get-Field $Facts 'asOf' $null))
  try{foreach($p in $patients){Set-P $p 'Items' (Get-PatientItems $p);Set-P $p 'Needs' @(Get-PatientNeeds $p);Set-P $p 'Urgency' (Get-PatientUrgency $p)};$nonRisk=@(Get-Field $Facts 'scorecardNonRisk' @()|Where-Object{$null -ne $_});$applied=(Resolve-ProviderFlags $Flags $patients $nonRisk).applied}finally{$script:AsOf=$prev}
 }
 return [ordered]@{active=@($patients|Where-Object{!$applied.ContainsKey([string]$_.MemberID)});flagged=$applied.Count}
}
function Get-FactsKpi($Facts,[object[]]$Flags){$a=Get-ActivePatients $Facts $Flags;return (New-KpiSummary @($a.active) $a.flagged)}   # same counts as the report's top section
function New-KpiTotal([object[]]$Rows){
 $t=[ordered]@{};foreach($c in $script:KpiColumns){$t[$c[0]]=0};$n=0
 foreach($r in @($Rows)){if($null -eq $r){continue};$k=Get-Field $r 'kpi' $null;if($null -eq $k){continue};$n++;foreach($c in $script:KpiColumns){$t[$c[0]]+=[int](Get-Field $k $c[0] 0)}}
 return [ordered]@{providers=$n;kpi=$t}
}
$script:GroupSort=@(@{Expression={([string]$_).StartsWith('(')}},@{Expression={[string]$_}})   # named groups alphabetically, then '(No location)' / '(No risk pool)'
$script:OverviewLastSave=[DateTime]::MinValue
function Set-RunProgress($Run,[string]$Path,[int]$Percent,[string]$Stage){
 # Progress of an overview or list worker is written at most once a second (a large pool counts many providers per second); the start and the final stages always are.
 $Run.state='Running';$Run.percent=$Percent;$Run.stage=$Stage
 if($Percent -gt 5 -and $Percent -lt 94 -and ([DateTime]::UtcNow-$script:OverviewLastSave).TotalMilliseconds -lt 1000){return}
 Save-JsonAtomic $Path $Run;$script:OverviewLastSave=[DateTime]::UtcNow
}
function Get-ChosenProfiles([object[]]$Npis){
 $byNpi=@{};foreach($p in @(Get-ProviderProfiles)){if($null -ne $p){$byNpi[[string](Get-P $p 'npi' '')]=$p}}
 $chosen=@();foreach($n in @($Npis)){$n=[string]$n;if($byNpi.ContainsKey($n)){$chosen+=$byNpi[$n]}}
 if($chosen.Count -eq 0){throw 'None of the chosen providers has a saved profile any more.'}
 return @($chosen)
}
function Read-GroupSources([object[]]$Chosen,$Run,[string]$Path,[int]$From,[int]$Span){
 # Reads each source once for every chosen provider and groups its rows by the source's provider column (progress runs From..From+Span).
 $groups=@{};$read=@();$step=0;$status=@{};foreach($s in @(Status)){$status[[string]$s.sourceKey]=$s}
 foreach($key in $script:OverviewSources){
  $step++;$set=@{};foreach($p in $Chosen){$al=Get-OverviewAlias (Get-P $p 'aliases' $null) $key;if($al){$set[$al]=$true}}
  if($set.Count -eq 0){continue}
  $fieldKey=$(if($key -like 'HR-*'){'HR'}else{$key});$filter=[string]$script:FactFilters[$fieldKey];$name=$(if($status.ContainsKey($key)){[string]$status[$key].displayName}else{$key})
  Set-RunProgress $Run $Path ([int]($From+$Span*($step-1)/$script:OverviewSources.Count)) ('Reading '+$name+' for '+$Chosen.Count+' provider'+$(if($Chosen.Count -eq 1){''}else{'s'}))
  $rows=@(Get-SourceRows $key $script:FactFields[$fieldKey] $filter '' -FilterSet $set)
  $g=@{};foreach($r in $rows){$v=([string]$r.$filter).Trim();if(!$g.ContainsKey($v)){$g[$v]=New-Object Collections.Generic.List[object]};$g[$v].Add($r)}
  $groups[$key]=$g
  $st=$(if($status.ContainsKey($key)){$status[$key]}else{$null});$read+=[ordered]@{sourceKey=$key;displayName=$name;importedUtc=$(if($st){ConvertTo-IsoText (Get-Field $st 'importedUtc' '')}else{''});lastWriteTime=$(if($st){ConvertTo-IsoText (Get-Field $st 'lastWriteTime' '')}else{''})}
 }
 return [ordered]@{groups=$groups;sources=@($read)}
}
function Get-ProviderFactsFromGroups($Groups,$P){
 # One provider's patient facts from rows read by Read-GroupSources, joined exactly as a single report joins them.
 $a=Get-P $P 'aliases' $null;$exportAlias=Get-OverviewAlias $a 'Export';$hrKey=Get-HrSourceKey $a
 $exports=@(Get-GroupedRows $Groups 'Export' $exportAlias)
 $facts=New-PatientFacts $exports @(Get-GroupedRows $Groups 'PtListQuality' (Get-OverviewAlias $a 'PtListQuality')) @(Get-GroupedRows $Groups 'DiabetesScorecard' (Get-OverviewAlias $a 'DiabetesScorecard')) @(Get-GroupedRows $Groups 'SerialScheduling' (Get-OverviewAlias $a 'SerialScheduling')) @(Get-GroupedRows $Groups 'RiskPopulationOutreach' (Get-OverviewAlias $a 'RiskPopulationOutreach')) @($(if($hrKey){Get-GroupedRows $Groups $hrKey (Get-OverviewAlias $a $hrKey)}))
 $note=$(if($exports.Count -eq 0){'No Export rows for "'+$exportAlias+'"; the name may have changed in the latest export. Check the profile in the Provider Wizard.'}else{''})
 return [ordered]@{facts=$facts;note=$note}
}
function Invoke-OverviewRun([string]$Id){
 # Runs in its own worker process: reads each source once for every selected provider, then counts per provider exactly as a report would.
 $path=Get-OverviewPath $Id;$ov=Json $path;if(!$ov){throw 'Overview was not found.'}
 $sw=[Diagnostics.Stopwatch]::StartNew()
 try{
  Set-P $ov 'workerPid' $PID;Set-P $ov 'startedUtc' ([DateTime]::UtcNow.ToString('o'));Set-RunProgress $ov $path 3 'Reading provider profiles'
  $chosen=@(Get-ChosenProfiles @($ov.npis))
  $src=Read-GroupSources $chosen $ov $path 5 60;$groups=$src.groups;$read=@($src.sources)
  $rowsOut=@();$i=0
  foreach($p in $chosen){
   $i++;$a=Get-P $p 'aliases' $null;$npi=[string](Get-P $p 'npi' '');$display=[string](Get-P $p 'displayName' $npi)
   $loc=([string](Get-P $p 'location' '')).Trim();$pool=([string](Get-P $p 'riskPool' '')).Trim()
   Set-RunProgress $ov $path ([int](65+28*($i-1)/$chosen.Count)) ('Counting patients for '+$display+' ('+$i+' of '+$chosen.Count+')')
   $row=[ordered]@{npi=$npi;displayName=$display;location=$(if($loc){$loc}else{'(No location)'});riskPool=$(if($pool){$pool}else{'(No risk pool)'});kpi=$null;note='';error=''}
   try{
    $pf=Get-ProviderFactsFromGroups $groups $p
    $row.kpi=Get-FactsKpi $pf.facts @(Get-ProviderFlags $npi);$row.note=$pf.note
   }catch{$row.error=$_.Exception.Message;Log 'OVERVIEW_PROVIDER' 'FAILED' ($npi+' '+$_.Exception.Message)}
   $rowsOut+=$row
  }
  $locations=@();foreach($ln in @($rowsOut|ForEach-Object{$_.location}|Select-Object -Unique|Sort-Object -Property $script:GroupSort)){$members=@($rowsOut|Where-Object{$_.location -eq $ln}|Sort-Object -Property @{Expression={$_.displayName}});$locations+=[ordered]@{location=$ln;rows=$members;total=(New-KpiTotal $members)}}
  $pools=@();foreach($pn in @($rowsOut|ForEach-Object{$_.riskPool}|Select-Object -Unique|Sort-Object -Property $script:GroupSort)){$t=New-KpiTotal @($rowsOut|Where-Object{$_.riskPool -eq $pn});$pools+=[ordered]@{riskPool=$pn;providers=$t.providers;kpi=$t.kpi}}
  Set-P $ov 'locations' $locations;Set-P $ov 'pools' $pools;Set-P $ov 'total' (New-KpiTotal $rowsOut);Set-P $ov 'sources' $read;Set-P $ov 'asOf' ([DateTime]::Now.ToString('o'));Set-P $ov 'failedProviders' @($rowsOut|Where-Object{$_.error}).Count
  Set-RunProgress $ov $path 94 'Rendering PDF'
  $ov=Json $path;$stageHtml=Join-Path $script:Paths.Staging ('overview-'+$Id+'.html')
  try{
   [IO.File]::WriteAllText($stageHtml,(ConvertTo-OverviewDocument $ov),(New-Object Text.UTF8Encoding($false)))
   # Rendered under a space-free staging name (the browser command line is not quoted), then moved to its readable name.
   $stagePdf=Join-Path $script:Paths.Staging ('overview-'+$Id+'.pdf');Convert-HtmlPdf $stageHtml $stagePdf
   $built=ConvertTo-DateValue (Get-P $ov 'asOf' $null);if(!$built){$built=Get-Date};$pdf=Join-Path $script:Paths.Overviews ('Provider Overview - '+$built.ToString('yyyy-MM-dd HHmmss')+'.pdf');Move-Item -LiteralPath $stagePdf -Destination $pdf -Force;Set-P $ov 'pdfPath' $pdf
  }catch{Set-P $ov 'pdfError' $_.Exception.Message;Log 'OVERVIEW_PDF' 'WARN' $_.Exception.Message}finally{Remove-Item -LiteralPath $stageHtml,(Join-Path $script:Paths.Staging ('overview-'+$Id+'.pdf')) -Force -ErrorAction SilentlyContinue}
  $ov.state='Completed';$ov.percent=100;$ov.stage='Overview complete';Set-P $ov 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $ov 'seconds' ([Math]::Round($sw.Elapsed.TotalSeconds,1));Save-JsonAtomic $path $ov
  Log 'OVERVIEW_COMPLETED' 'OK' ($chosen.Count.ToString()+' providers in '+[Math]::Round($sw.Elapsed.TotalSeconds,1)+'s')
 }catch{$msg=$_.Exception.Message;$cur=Json $path;if(!$cur){$cur=$ov};$cur.state='Failed';$cur.stage='Overview failed';Set-P $cur 'errorSummary' $msg;Set-P $cur 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $cur;Log 'OVERVIEW_FAILED' 'FAILED' $msg}
}
function Start-RunWorker([string]$Switch,[string]$Id,[string]$LogName){
 # A background copy of this script for one overview or list run; its output goes to logs\<LogName>.out.log / .err.log.
 $exe=(Get-Process -Id $PID).Path
 $opt=@{FilePath=$exe;ArgumentList=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$PSCommandPath+'"'),('-'+$Switch),$Id,'-NoBrowser');RedirectStandardOutput=(Join-Path $script:Paths.Logs ($LogName+'.out.log'));RedirectStandardError=(Join-Path $script:Paths.Logs ($LogName+'.err.log'));PassThru=$true}
 if([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT){$opt['WindowStyle']='Hidden'}
 return (Start-Process @opt)
}
function Start-OverviewWorker([string]$Id){return (Start-RunWorker 'RunOverviewId' $Id ('overview-'+$Id))}
function Get-Overviews{return @(Get-ChildItem $script:Paths.State -Filter 'overview-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Where-Object{$null -ne $_}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'requestedUtc' '')}} -Descending)}
function Get-RunStallReason($Run,[string]$Id,[string]$Noun,[string]$Again,[string]$LogName){
 # A worker that died (server or computer restarted) would otherwise leave its run "Running" forever; returns why it stopped, or ''.
 $state=[string](Get-P $Run 'state' '');if($state -ne 'Starting' -and $state -ne 'Running'){return ''}
 $workerPid=[int](Get-P $Run 'workerPid' 0)
 if($workerPid -gt 0){if(!(Test-WorkerAlive ([pscustomobject]@{workerPid=$workerPid;jobId=$Id}))){return ('The '+$Noun+' worker stopped before finishing (the server or computer may have restarted). '+$Again)};return ''}
 $req=ConvertTo-DateValue (Get-P $Run 'requestedUtc' $null);if($req -and ([DateTime]::UtcNow-$req.ToUniversalTime()).TotalSeconds -gt 120){return ('The '+$Noun+' worker did not start; see logs\'+$LogName+'.err.log.')}
 return ''
}
function Update-OverviewLiveness($Ov){
 $id=[string]$Ov.overviewId;$reason=Get-RunStallReason $Ov $id 'overview' 'Build the overview again.' ('overview-'+$id)
 if(!$reason){return $false}
 $Ov.state='Failed';$Ov.stage='Overview failed';Set-P $Ov 'errorSummary' $reason;Save-JsonAtomic (Get-OverviewPath ([string]$Ov.overviewId)) $Ov;Log 'OVERVIEW_INTERRUPTED' 'WARN' ([string]$Ov.overviewId);return $true
}
function Get-OverviewModel([string]$Id,[switch]$Summary){
 $ov=Json (Get-OverviewPath $Id);if(!$ov){throw 'Overview was not found.'}
 $null=Update-OverviewLiveness $ov
 $m=[ordered]@{overviewId=[string]$ov.overviewId;state=[string]$ov.state;percent=[int](Get-P $ov 'percent' 0);stage=[string](Get-P $ov 'stage' '');requestedUtc=(ConvertTo-IsoText (Get-P $ov 'requestedUtc' ''));completedUtc=(ConvertTo-IsoText (Get-P $ov 'completedUtc' ''));asOf=(ConvertTo-IsoText (Get-P $ov 'asOf' ''));seconds=(Get-P $ov 'seconds' $null);npis=@(Get-P $ov 'npis' @());errorSummary=[string](Get-P $ov 'errorSummary' '');pdfError=[string](Get-P $ov 'pdfError' '');failedProviders=[int](Get-P $ov 'failedProviders' 0);providerCount=@(Get-P $ov 'npis' @()).Count;selection=(Get-OverviewSelectionText $ov);totalPatients=$null;tableHtml='';sourcesText='';xlsxUrl=$null;pdfUrl=$null}
 if($m.state -eq 'Completed'){
  $m.totalPatients=[int](Get-Field (Get-Field (Get-P $ov 'total' $null) 'kpi' $null) 'total' 0);$m.xlsxUrl='/overview.xlsx?id='+$m.overviewId
  if(!$Summary){$m.tableHtml=ConvertTo-OverviewTablesHtml $ov;$m.sourcesText=Get-OverviewSourcesText $ov}
  $pdf=[string](Get-P $ov 'pdfPath' '');if($pdf -and (Test-Path -LiteralPath $pdf)){$m.pdfUrl='/overview.pdf?id='+$m.overviewId}elseif($pdf -and !$m.pdfError){$m.pdfError='The PDF was moved or deleted from the overviews folder.'}
 }
 return $m
}
function Get-OverviewPageModel{
 $all=@(Get-Overviews);$history=@(foreach($o in $all){if([string](Get-P $o 'state' '') -in @('Completed','Failed')){Get-OverviewModel ([string]$o.overviewId) -Summary}})
 return [ordered]@{providers=@(Get-OverviewProviders);poc=@(Get-PocLatestByLocation);latest=$(if($all.Count){Get-OverviewModel ([string]$all[0].overviewId)}else{$null});history=$history}
}
function New-SelectionSummary([object[]]$All,[hashtable]$Chosen,[string]$Field){
 # Per risk pool or location: how many of its saved profiles were picked (only groups with at least one pick).
 $out=@();foreach($name in @($All|ForEach-Object{$_[$Field]}|Select-Object -Unique|Sort-Object -Property $script:GroupSort)){$members=@($All|Where-Object{$_[$Field] -eq $name});$picked=@($members|Where-Object{$Chosen.ContainsKey($_.npi)}).Count;if($picked -gt 0){$out+=[ordered]@{name=$name;selected=$picked;total=$members.Count}}}
 return @($out)
}
function Get-OverviewSelectionText($Ov){
 # "CRYSTAL RUN (all 12), PROHEALTH (2 of 8)"; overviews built before the selection was recorded fall back to their result groups.
 $fmt={param($Items)(@(foreach($i in @($Items)){if($null -eq $i){continue};$k=[int](Get-Field $i 'selected' 0);$t=[int](Get-Field $i 'total' 0);[string](Get-Field $i 'name' '')+$(if($t -le 0){' ('+$k+')'}elseif($k -ge $t){' (all '+$t+')'}else{' ('+$k+' of '+$t+')'})}) -join ', ')}
 $sel=Get-P $Ov 'selection' $null
 if($sel){return [ordered]@{riskPools=(& $fmt (Get-Field $sel 'riskPools' @()));locations=(& $fmt (Get-Field $sel 'locations' @()))}}
 $pools=@(foreach($p in @(Get-P $Ov 'pools' @())){if($p){[ordered]@{name=[string](Get-Field $p 'riskPool' '');selected=[int](Get-Field $p 'providers' 0);total=0}}})
 $locs=@(foreach($l in @(Get-P $Ov 'locations' @())){if($l){[ordered]@{name=[string](Get-Field $l 'location' '');selected=@(Get-Field $l 'rows' @()).Count;total=0}}})
 return [ordered]@{riskPools=(& $fmt $pools);locations=(& $fmt $locs)}
}
function Get-RequestedProviders($Body){
 # The picked NPIs (each must have a saved profile) and, per risk pool and location, how many of its providers were picked.
 $npis=@(@(Get-P $Body 'npis' @())|ForEach-Object{([string]$_).Trim()}|Where-Object{$_ -match '^\d{10}$'}|Select-Object -Unique)
 if($npis.Count -eq 0){throw 'Choose at least one provider.'}
 $all=@(Get-OverviewProviders);$known=@{};foreach($p in $all){$known[$p.npi]=$true}
 $missing=@($npis|Where-Object{!$known.ContainsKey($_)});if($missing.Count -gt 0){throw ('No saved profile for NPI '+($missing -join ', ')+'; reload the page.')}
 $chosen=@{};foreach($n in $npis){$chosen[$n]=$true}
 return [ordered]@{npis=@($npis);selection=[ordered]@{riskPools=@(New-SelectionSummary $all $chosen 'riskPool');locations=@(New-SelectionSummary $all $chosen 'location')}}
}
function New-OverviewRun($Body){
 $req=Get-RequestedProviders $Body;$npis=@($req.npis);$selection=$req.selection
 foreach($o in @(Get-Overviews)){if([string](Get-P $o 'state' '') -in @('Starting','Running') -and !(Update-OverviewLiveness $o)){throw 'An overview is already being built; wait for it to finish.'}}
 $id=[Guid]::NewGuid().ToString('N');$path=Get-OverviewPath $id
 $ov=[ordered]@{overviewVersion=1;overviewId=$id;state='Starting';percent=1;stage='Launching overview worker';requestedUtc=[DateTime]::UtcNow.ToString('o');npis=$npis;selection=$selection;workerPid=0}
 Save-JsonAtomic $path $ov;Log 'OVERVIEW_REQUESTED' 'OK' ($npis.Count.ToString()+' providers')
 if($script:OverviewInline){Invoke-OverviewRun $id}
 else{try{$null=Start-OverviewWorker $id}catch{$ov.state='Failed';$ov.stage='Worker launch failed';$ov['errorSummary']=$_.Exception.Message;Save-JsonAtomic $path $ov;Log 'OVERVIEW_FAILED' 'FAILED' $_.Exception.Message}}
 foreach($old in @(Get-Overviews|Select-Object -Skip $script:OverviewKeep)){$oid=[string](Get-P $old 'overviewId' '');if($oid -match '^[a-f0-9]{32}$'){Remove-Item -LiteralPath (Get-OverviewPath $oid) -Force -ErrorAction SilentlyContinue;Remove-Item -LiteralPath (Join-Path $script:Paths.Logs ('overview-'+$oid+'.out.log')),(Join-Path $script:Paths.Logs ('overview-'+$oid+'.err.log')) -Force -ErrorAction SilentlyContinue}}
 return (Get-OverviewModel $id)
}
function Get-OverviewSourcesText($Ov){
 $parts=@();foreach($s in @(Get-P $Ov 'sources' @())){if($null -eq $s){continue};$d=ConvertTo-DateValue (Get-Field $s 'importedUtc' '');if(!$d){$d=ConvertTo-DateValue (Get-Field $s 'lastWriteTime' '')};$parts+=([string](Get-Field $s 'displayName' '')+' '+$(if($d){$d.ToLocalTime().ToString('M/d/yyyy')}else{'(date unknown)'}))}
 return ('Source files: '+($parts -join ', '))
}
function Get-KpiCellsHtml($Kpi){$h='';foreach($c in $script:KpiColumns){$h+='<td class="n">'+$(if($null -eq $Kpi){'-'}else{[string](Get-Field $Kpi $c[0] 0)})+'</td>'};return $h}
function ConvertTo-OverviewTablesHtml($Ov){
 # One table: providers grouped by location with a subtotal under each, grand total last; then totals by risk pool. Shared by the page and the PDF.
 $heads='';foreach($c in $script:KpiColumns){$heads+='<th class="n">'+(ConvertTo-HtmlEncoded $c[1])+'</th>'};$span=2+$script:KpiColumns.Count
 $body=''
 foreach($l in @(Get-P $Ov 'locations' @())){
  if($null -eq $l){continue};$members=@(Get-Field $l 'rows' @()|Where-Object{$null -ne $_});$lt=Get-Field $l 'total' $null
  $body+='<tr class="loc"><th colspan="'+$span+'">'+(ConvertTo-HtmlEncoded (Get-Field $l 'location' ''))+'</th></tr>'
  foreach($r in $members){$note=[string](Get-Field $r 'note' '');$err=[string](Get-Field $r 'error' '');$body+='<tr><td>'+(ConvertTo-HtmlEncoded (Get-Field $r 'displayName' ''))+$(if($err){'<div class="rowerr">Not counted: '+(ConvertTo-HtmlEncoded $err)+'</div>'}elseif($note){'<div class="rownote">'+(ConvertTo-HtmlEncoded $note)+'</div>'}else{''})+'</td><td class="pool">'+(ConvertTo-HtmlEncoded (Get-Field $r 'riskPool' ''))+'</td>'+(Get-KpiCellsHtml (Get-Field $r 'kpi' $null))+'</tr>'}
  $body+='<tr class="sub"><td>'+(ConvertTo-HtmlEncoded (Get-Field $l 'location' ''))+' total</td><td>'+[int](Get-Field $lt 'providers' 0)+' provider'+$(if([int](Get-Field $lt 'providers' 0) -eq 1){''}else{'s'})+'</td>'+(Get-KpiCellsHtml (Get-Field $lt 'kpi' $null))+'</tr>'
 }
 $t=Get-P $Ov 'total' $null;$tn=[int](Get-Field $t 'providers' 0)
 $body+='<tr class="grand"><td>All providers</td><td>'+$tn+' provider'+$(if($tn -eq 1){''}else{'s'})+'</td>'+(Get-KpiCellsHtml (Get-Field $t 'kpi' $null))+'</tr>'
 $poolRows='';foreach($p in @(Get-P $Ov 'pools' @())){if($null -eq $p){continue};$poolRows+='<tr><td class="pool">'+(ConvertTo-HtmlEncoded (Get-Field $p 'riskPool' ''))+'</td><td class="n">'+[int](Get-Field $p 'providers' 0)+'</td>'+(Get-KpiCellsHtml (Get-Field $p 'kpi' $null))+'</tr>'}
 $poolRows+='<tr class="grand"><td>All risk pools</td><td class="n">'+$tn+'</td>'+(Get-KpiCellsHtml (Get-Field $t 'kpi' $null))+'</tr>'
 return '<table class="ov"><thead><tr><th>Provider</th><th>Risk pool</th>'+$heads+'</tr></thead><tbody>'+$body+'</tbody></table><h3 class="ovh">Totals by risk pool</h3><table class="ov"><thead><tr><th>Risk pool</th><th class="n">Providers</th>'+$heads+'</tr></thead><tbody>'+$poolRows+'</tbody></table>'
}
$script:OverviewCss='table.ov{border-collapse:collapse;width:100%;margin-top:8px}table.ov th,table.ov td{padding:5px 7px;border-bottom:1px solid #dce4ef;text-align:left;vertical-align:top}table.ov thead th{background:#17365d;color:#fff;font-weight:600;vertical-align:bottom}table.ov .n,table.ov .pool{white-space:nowrap}table.ov .n{text-align:right}table.ov thead th.n{white-space:normal;max-width:80px}table.ov tr.loc th{background:#eaf1f8;color:#17365d;font-size:1.05em;padding-top:9px}table.ov tr.sub td{background:#f4f7fb;font-weight:700;border-bottom:2px solid #c8d2e0}table.ov tr.grand td{background:#dfe9f5;font-weight:700;border-top:2px solid #17365d}.rownote{color:#9a5b00;font-size:.85em}.rowerr{color:#a61b1b;font-size:.85em}h3.ovh{color:#17365d;margin:18px 0 0}'
function ConvertTo-OverviewDocument($Ov){
 # Printable landscape page for the PDF.
 $n=@(Get-P $Ov 'npis' @()).Count;$asOf=ConvertTo-DateValue (Get-P $Ov 'asOf' $null);$when=$(if($asOf){$asOf.ToString('g')}else{''})
 return '<!doctype html><html><head><meta charset="utf-8"><title>Provider Overview</title><style>@page{size:letter landscape;margin:.4in}body{font:10px Segoe UI,Arial;color:#172033;margin:0}h1{font-size:17px;color:#17365d;margin:0 0 2px}.muted{color:#667085}table.ov{page-break-inside:auto}table.ov tr{page-break-inside:avoid}table.ov thead{display:table-header-group}'+$script:OverviewCss+'</style></head><body><h1>Provider Overview</h1><div class="muted">'+$n+' provider'+$(if($n -eq 1){''}else{'s'})+' &middot; generated '+(ConvertTo-HtmlEncoded $when)+' &middot; '+(ConvertTo-HtmlEncoded (Get-OverviewSourcesText $Ov))+' &middot; flagged patients are excluded from every count, as on each dashboard</div>'+(ConvertTo-OverviewTablesHtml $Ov)+'</body></html>'
}
function Send-Bytes($Context,[string]$Type,[byte[]]$Bytes,[string]$Disposition){$Context.Response.StatusCode=200;$Context.Response.ContentType=$Type;$Context.Response.Headers['Content-Disposition']=$Disposition;$Context.Response.Headers['Cache-Control']='no-store';$Context.Response.ContentLength64=$Bytes.Length;$Context.Response.OutputStream.Write($Bytes,0,$Bytes.Length);$Context.Response.Close()}
function Get-CompletedOverview([string]$Id){$ov=Json (Get-OverviewPath $Id);if(!$ov -or [string](Get-P $ov 'state' '') -ne 'Completed'){throw 'That overview is not complete.'};return $ov}
function New-OverviewWorkbook($Ov){
 # Same layout as the page: providers grouped by location with subtotals and a grand total, then totals by risk pool.
 try{Add-Type -AssemblyName System.Drawing -ErrorAction Stop}catch{}
 $pkg=New-Object OfficeOpenXml.ExcelPackage;$ws=$pkg.Workbook.Worksheets.Add('Overview');$kc=$script:KpiColumns.Count
 $asOf=ConvertTo-DateValue (Get-P $Ov 'asOf' $null)
 $ws.Cells[1,1].Value='Provider Overview';$ws.Cells[1,1].Style.Font.Bold=$true;$ws.Cells[1,1].Style.Font.Size=14
 $ws.Cells[2,1].Value=('Generated '+$(if($asOf){$asOf.ToString('g')}else{''})+'. '+(Get-OverviewSourcesText $Ov)+'. Flagged patients are excluded from every count, as on each dashboard.')
 $fill={param($Row,[int]$LastCol,[string]$Hex,[bool]$Bold)$rng=$ws.Cells[$Row,1,$Row,$LastCol];$rng.Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid;$h=$Hex.TrimStart('#');$rng.Style.Fill.BackgroundColor.SetColor([Drawing.Color]::FromArgb([Convert]::ToInt32($h.Substring(0,2),16),[Convert]::ToInt32($h.Substring(2,2),16),[Convert]::ToInt32($h.Substring(4,2),16)));if($Bold){$rng.Style.Font.Bold=$true}}
 $r=4;$heads=@('Location','Provider','NPI','Risk Pool')+@($script:KpiColumns|ForEach-Object{$_[1]});$last=$heads.Count
 for($c=0;$c -lt $last;$c++){$ws.Cells[$r,($c+1)].Value=$heads[$c]};& $fill $r $last '#17365D' $true;$ws.Cells[$r,1,$r,$last].Style.Font.Color.SetColor([Drawing.Color]::White);$ws.Cells[$r,1,$r,$last].Style.WrapText=$true;$headRow=$r
 $put={param($Row,[int]$Col,$Kpi)for($k=0;$k -lt $kc;$k++){if($null -ne $Kpi){$ws.Cells[$Row,($Col+$k)].Value=[int](Get-Field $Kpi $script:KpiColumns[$k][0] 0)}}}
 foreach($l in @(Get-P $Ov 'locations' @())){
  if($null -eq $l){continue};$ln=[string](Get-Field $l 'location' '')
  foreach($x in @(Get-Field $l 'rows' @()|Where-Object{$null -ne $_})){$r++;$ws.Cells[$r,1].Value=$ln;$ws.Cells[$r,2].Value=[string](Get-Field $x 'displayName' '');$ws.Cells[$r,3].Value=[string](Get-Field $x 'npi' '');$ws.Cells[$r,4].Value=[string](Get-Field $x 'riskPool' '');& $put $r 5 (Get-Field $x 'kpi' $null);$err=[string](Get-Field $x 'error' '');if($err){$ws.Cells[$r,($last+1)].Value=('Not counted: '+$err)}else{$note=[string](Get-Field $x 'note' '');if($note){$ws.Cells[$r,($last+1)].Value=$note}}}
  $lt=Get-Field $l 'total' $null;$r++;$ws.Cells[$r,1].Value=$ln;$ws.Cells[$r,2].Value=($ln+' total ('+[int](Get-Field $lt 'providers' 0)+')');& $put $r 5 (Get-Field $lt 'kpi' $null);& $fill $r $last '#F4F7FB' $true
 }
 $t=Get-P $Ov 'total' $null;$r++;$ws.Cells[$r,1].Value='All locations';$ws.Cells[$r,2].Value=('All providers ('+[int](Get-Field $t 'providers' 0)+')');& $put $r 5 (Get-Field $t 'kpi' $null);& $fill $r $last '#DFE9F5' $true
 $r+=3;$ws.Cells[$r,1].Value='Totals by risk pool';$ws.Cells[$r,1].Style.Font.Bold=$true;$r++
 $pheads=@('Risk Pool','Providers')+@($script:KpiColumns|ForEach-Object{$_[1]});$plast=$pheads.Count;for($c=0;$c -lt $plast;$c++){$ws.Cells[$r,($c+1)].Value=$pheads[$c]};& $fill $r $plast '#17365D' $true;$ws.Cells[$r,1,$r,$plast].Style.Font.Color.SetColor([Drawing.Color]::White);$ws.Cells[$r,1,$r,$plast].Style.WrapText=$true
 foreach($p in @(Get-P $Ov 'pools' @())){if($null -eq $p){continue};$r++;$ws.Cells[$r,1].Value=[string](Get-Field $p 'riskPool' '');$ws.Cells[$r,2].Value=[int](Get-Field $p 'providers' 0);& $put $r 3 (Get-Field $p 'kpi' $null)}
 $r++;$ws.Cells[$r,1].Value='All risk pools';$ws.Cells[$r,2].Value=[int](Get-Field $t 'providers' 0);& $put $r 3 (Get-Field $t 'kpi' $null);& $fill $r $plast '#DFE9F5' $true
 $ws.Column(1).Width=24;$ws.Column(2).Width=30;$ws.Column(3).Width=13;$ws.Column(4).Width=18;for($c=5;$c -le $last;$c++){$ws.Column($c).Width=12};$ws.Column($last+1).Width=60
 $ws.View.FreezePanes(($headRow+1),3)
 try{return ,$pkg.GetAsByteArray()}finally{$pkg.Dispose()}
}
function Send-OverviewFile($Context,[string]$Id,[string]$Type){
 $ov=Get-CompletedOverview $Id;$asOf=ConvertTo-DateValue (Get-P $ov 'asOf' $null);$stamp=$(if($asOf){$asOf.ToString('yyyy-MM-dd HHmmss')}else{Get-Date -Format 'yyyy-MM-dd HHmmss'})
 if($Type -eq 'xlsx'){Send-Bytes $Context 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' (New-OverviewWorkbook $ov) ('attachment; filename="Provider Overview - '+$stamp+'.xlsx"');return}
 $pdf=[string](Get-P $ov 'pdfPath' '');$full=$(if($pdf){[IO.Path]::GetFullPath($pdf)}else{''});$root=[IO.Path]::GetFullPath($script:Paths.Overviews)
 if(!$full -or !$full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or !(Test-Path -LiteralPath $full)){throw 'The PDF for this overview is not available.'}
 Send-Bytes $Context 'application/pdf' ([IO.File]::ReadAllBytes($full)) ('inline; filename="'+[IO.Path]::GetFileName($full)+'"')
}
# Shared by the Overview and New Patients pages: provider picker (risk pool / location chips, providers by location), progress bar and history table.
$script:RunPageCss='.chips{display:flex;flex-wrap:wrap;gap:6px;margin:4px 0 10px}.chip{border:1px solid #c8d2e0;border-radius:16px;padding:4px 10px;background:#f8fafc;cursor:pointer;user-select:none}.chip input{margin:0 5px 0 0;vertical-align:-1px}.grp{font-weight:600;color:#17365d;margin-top:10px}.provs{columns:3 260px;column-gap:18px;margin-top:6px}.locgrp{break-inside:avoid;margin:0 0 10px}.lochead{font-weight:600;color:#17365d;border-bottom:1px solid #dce4ef;margin-bottom:3px}.prov{display:block;padding:2px 0;cursor:pointer}.prov input{margin:0 6px 0 0}.progress{height:10px;background:#e4e9f0;border-radius:6px;overflow:hidden;margin:6px 0;max-width:600px}.progress span{display:block;height:100%;background:#1769aa;transition:width .3s}.scroll{overflow-x:auto}table.hist{border-collapse:collapse;width:100%}table.hist th,table.hist td{padding:6px 8px;border-bottom:1px solid #dce4ef;text-align:left;vertical-align:top}table.hist th{background:#eaf1f8}table.hist .n{text-align:right}table.hist tr.showing td{background:#f4f9ff}'
$script:PickerJs=@'
const el=id=>document.getElementById(id);
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function api(u,o){const r=await fetch(u,o);const text=await r.text();let j=null;try{j=text?JSON.parse(text):null}catch(e){throw Error('Server returned invalid JSON: '+text.slice(0,200))}if(!r.ok)throw Error((j&&j.error)||('Request failed ('+r.status+')'));return j}
function when(iso){if(!iso)return '';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleString()}
let providers=[],sel=new Set();
function sortKey(a,b){const x=(a.startsWith('(')?'~':'')+a.toLowerCase(),y=(b.startsWith('(')?'~':'')+b.toLowerCase();return x<y?-1:x>y?1:0}
function groups(field){const g={};providers.forEach(p=>{(g[p[field]]=g[p[field]]||[]).push(p)});return Object.keys(g).sort(sortKey).map(k=>({name:k,list:g[k].sort((a,b)=>sortKey(a.displayName,b.displayName))}))}
function chips(field){return '<div class="chips">'+groups(field).map(g=>'<label class="chip"><input type="checkbox" data-field="'+field+'" data-name="'+esc(g.name)+'">'+esc(g.name)+' <span class="muted">('+g.list.length+')</span></label>').join('')+'</div>'}
function renderPicker(){
 if(!providers.length){el('picker').innerHTML='<span class="muted">No saved provider profiles yet. Map providers in the Provider Wizard and save their profiles first.</span>';syncChecks();return}
 el('picker').innerHTML='<div class="grp">Risk pools</div>'+chips('riskPool')+'<div class="grp">Locations</div>'+chips('location')+'<div class="grp">Providers</div><div class="provs">'+groups('location').map(g=>'<div class="locgrp"><div class="lochead">'+esc(g.name)+'</div>'+g.list.map(p=>'<label class="prov"><input type="checkbox" data-npi="'+esc(p.npi)+'">'+esc(p.displayName)+' <span class="muted">'+esc(p.riskPool)+'</span></label>').join('')+'</div>').join('')+'</div>';
 syncChecks();
}
function members(field,name){return providers.filter(p=>p[field]===name).map(p=>p.npi)}
function pickerSync(){
 el('picker').querySelectorAll('input[data-npi]').forEach(cb=>{cb.checked=sel.has(cb.dataset.npi)});
 el('picker').querySelectorAll('input[data-field]').forEach(cb=>{const m=members(cb.dataset.field,cb.dataset.name);const n=m.filter(x=>sel.has(x)).length;cb.checked=n>0&&n===m.length;cb.indeterminate=n>0&&n<m.length});
 el('count').textContent=sel.size?sel.size+' of '+providers.length+' provider'+(providers.length===1?'':'s')+' selected':'No providers selected';
 pocLinksSync();
}
el('picker').addEventListener('change',e=>{const cb=e.target;if(cb.dataset.npi){if(cb.checked)sel.add(cb.dataset.npi);else sel.delete(cb.dataset.npi)}else if(cb.dataset.field){members(cb.dataset.field,cb.dataset.name).forEach(n=>{if(cb.checked)sel.add(n);else sel.delete(n)})}syncChecks()});
let pocReports=[];
function presetLocation(){const q=new URLSearchParams(window.location.search).get('location');if(!q)return false;const m=providers.filter(p=>p.location.toLowerCase()===q.toLowerCase());if(!m.length)return false;sel.clear();m.forEach(p=>sel.add(p.npi));return true}
function pocLinksSync(){const box=el('pocLinks');if(!box)return;const locs=new Set(providers.filter(p=>sel.has(p.npi)).map(p=>p.location.toLowerCase()));const list=pocReports.filter(r=>locs.has(r.location.toLowerCase()));box.style.display=list.length?'':'none';box.innerHTML=list.length?'POC coding report'+(list.length===1?'':'s')+' for the selected location'+(list.length===1?'':'s')+': '+list.map(r=>'<a target="_blank" href="'+esc(r.url)+'" title="'+esc(r.department)+'">'+esc(r.location)+'</a> <span class="muted">('+esc(when(r.completedUtc))+')</span>').join(' &middot; '):''}
el('all').onclick=()=>{providers.forEach(p=>sel.add(p.npi));syncChecks()};
el('none').onclick=()=>{sel.clear();syncChecks()};
'@
function OverviewPage{
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>Overview</title>
<style>
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px;max-width:1500px}
__TABCSS__
.card{background:white;border:1px solid #dce4ef;border-radius:8px;padding:18px;margin:14px 0}.card h2{margin:0 0 8px;font-size:17px;color:#17365d}.muted{color:#667085}.error{color:#a61b1b}.warn{color:#9a5b00}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:8px 0}
button,a.btn{padding:7px 12px;cursor:pointer;background:#1769aa;color:#fff;border:0;border-radius:5px;font:inherit;text-decoration:none;display:inline-block}button:disabled{opacity:.5;cursor:default}button.alt{background:#e4e9f0;color:#172033}button.go{background:#1d7a3a;font-weight:600;padding:10px 16px}
__RUNCSS__
__OVCSS__
</style></head><body>
<header><h1>Overview</h1><p>Draft 5.5 - the dashboard's top-section counts for a group of providers: one row per provider, subtotals by location and risk pool, grand total at the bottom</p></header>__NAV__
<main>
<section class="card"><h2>Choose providers</h2>
<p class="muted">Tick a risk pool or location to select all of its providers, or tick providers individually. Locations and risk pools come from each saved provider profile.</p>
<div id="picker"><span class="muted">Loading profiles...</span></div>
<div class="row"><button class="alt" id="all">Select all</button><button class="alt" id="none">Clear</button><button class="go" id="build" disabled>Build overview</button><span id="count" class="muted"></span></div>
<p id="pocLinks" style="display:none"></p>
<p class="muted">Counts are computed from the current source files in one pass (no reports or PDFs are generated per provider), with flagged patients excluded exactly as on each dashboard.</p>
</section>
<section class="card"><h2>Overview</h2><div id="result"><span class="muted">Loading...</span></div></section>
<section class="card"><h2>Previous overviews</h2><p class="muted">The last 30 overviews, newest first. View shows one again and restores its provider selection; its Excel and PDF are kept as built.</p><div id="history"><span class="muted">Loading...</span></div></section>
</main>
<script>
__PICKERJS__
let current=null,timer=null;
function busy(){return !!current&&(current.state==='Starting'||current.state==='Running')}
function syncChecks(){
 pickerSync();
 el('build').disabled=!sel.size||busy();el('build').textContent=busy()?'Building...':'Build overview';
 const note=el('selNote');if(note)note.style.display=selectionChanged()?'':'none';
}
function selectionChanged(){if(!current||current.state!=='Completed')return false;const a=[...sel].sort().join(','),b=[...current.npis].sort().join(',');return a!==b}
function showResult(o){
 current=o;
 if(!o){el('result').innerHTML='<span class="muted">No overview built yet. Choose providers above and select Build overview.</span>';syncChecks();return}
 if(o.state==='Starting'||o.state==='Running'){el('result').innerHTML='<p>Building the overview for '+o.providerCount+' provider'+(o.providerCount===1?'':'s')+'... '+o.percent+'%</p><div class="progress"><span style="width:'+Math.max(0,Math.min(100,o.percent))+'%"></span></div><span class="muted">'+esc(o.stage)+'</span>';syncChecks();return}
 if(o.state==='Failed'){el('result').innerHTML='<p class="error">The overview could not be built: '+esc(o.errorSummary)+'</p><span class="muted">Requested '+esc(when(o.requestedUtc))+'</span>';syncChecks();return}
 el('result').innerHTML='<div class="row"><b>'+o.providerCount+' provider'+(o.providerCount===1?'':'s')+'</b><span class="muted">built '+esc(when(o.completedUtc))+(o.seconds!=null?' in '+o.seconds+' s':'')+'</span><a class="btn" id="xlsx" href="'+esc(o.xlsxUrl)+'">Download Excel</a>'+(o.pdfUrl?'<a class="btn" id="pdf" target="_blank" href="'+esc(o.pdfUrl)+'">Open PDF</a>':'<span class="warn">PDF not available'+(o.pdfError?': '+esc(o.pdfError):'')+'</span>')+'</div><p class="muted">'+esc(o.sourcesText)+'. Flagged patients are excluded from every count, as on each dashboard.'+(o.failedProviders?' <span class="error">'+o.failedProviders+' provider'+(o.failedProviders===1?' was':'s were')+' not counted; see the red notes below.</span>':'')+'</p><p class="warn" id="selNote" style="display:none">Your selection has changed since this overview was built; select Build overview to update it.</p><div class="scroll">'+o.tableHtml+'</div>';
 syncChecks();
}
function historyHtml(list){
 if(!list.length)return '<span class="muted">None yet.</span>';
 return '<table class="hist"><thead><tr><th>Built</th><th>Providers</th><th>Risk pools</th><th>Locations</th><th class="n">Patients</th><th></th></tr></thead><tbody>'+list.map(h=>{const showing=current&&current.overviewId===h.overviewId;const s=h.selection||{};return '<tr'+(showing?' class="showing"':'')+'><td>'+esc(when(h.completedUtc||h.requestedUtc))+(showing?' <span class="muted">(shown above)</span>':'')+'</td><td>'+h.providerCount+'</td><td>'+esc(s.riskPools||'')+'</td><td>'+esc(s.locations||'')+'</td><td class="n">'+(h.totalPatients!=null?h.totalPatients:'')+'</td><td>'+(h.state==='Completed'?'<a href="#" data-view="'+esc(h.overviewId)+'">View</a> &middot; <a href="'+esc(h.xlsxUrl)+'">Excel</a>'+(h.pdfUrl?' &middot; <a target="_blank" href="'+esc(h.pdfUrl)+'">PDF</a>':''):'<span class="error">Failed: '+esc(h.errorSummary)+'</span>')+'</td></tr>'}).join('')+'</tbody></table>';
}
let history=[];
function renderHistory(){el('history').innerHTML=historyHtml(history)}
async function loadHistory(){try{const m=await api('/api/overview');history=m.history;renderHistory()}catch(e){el('history').innerHTML='<p class="error">'+esc(e.message)+'</p>'}}
el('history').addEventListener('click',async e=>{const a=e.target.closest('a[data-view]');if(!a)return;e.preventDefault();if(busy())return;try{const o=await api('/api/overview-run?id='+encodeURIComponent(a.dataset.view));sel.clear();o.npis.forEach(n=>{if(providers.some(p=>p.npi===n))sel.add(n)});showResult(o);renderHistory();el('result').scrollIntoView({behavior:'smooth'})}catch(err){el('result').innerHTML='<p class="error">'+esc(err.message)+'</p>'}});
let pollErrors=0;
function startPoll(id){if(timer)clearInterval(timer);pollErrors=0;timer=setInterval(async()=>{try{const o=await api('/api/overview-run?id='+encodeURIComponent(id));pollErrors=0;showResult(o);if(!busy()){clearInterval(timer);timer=null;loadHistory()}}catch(e){if(++pollErrors<5)return;clearInterval(timer);timer=null;el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>'}},1500)}
el('build').onclick=async()=>{if(!sel.size||busy())return;el('build').disabled=true;try{const o=await api('/api/overview',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({npis:[...sel]})});showResult(o);if(busy())startPoll(o.overviewId)}catch(e){el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>';syncChecks()}};
(async()=>{try{const m=await api('/api/overview');providers=m.providers;pocReports=m.poc||[];const last=m.latest;if(!presetLocation()&&last)last.npis.forEach(n=>{if(providers.some(p=>p.npi===n))sel.add(n)});renderPicker();showResult(last);history=m.history;renderHistory();if(busy())startPoll(last.overviewId)}catch(e){el('picker').innerHTML='<p class="error">'+esc(e.message)+'</p>'}})();
</script></body></html>
'@
 return $html.Replace('__NAV__',(Get-NavHtml 'overview')).Replace('__TABCSS__',$script:TabCss).Replace('__OVCSS__',$script:OverviewCss).Replace('__RUNCSS__',$script:RunPageCss).Replace('__PICKERJS__',$script:PickerJs)
}
# --- Draft 5.6: New Patients tab - per-provider lists of patients marked New Patient with no PCP visit in the last 12 months ---
$script:NewListKeep=30;$script:NewListInline=$false;$script:NewListMonths=12;$script:NewListTitle='New Patients Not Seen'
$script:NewListFormats=[ordered]@{pdf='PDF';html='HTML';xlsx='Excel';csv='CSV'}
$script:NewListGroupings=[ordered]@{combined='Combined';provider='Per provider';both='Combined and per provider'}
$script:NewListColumns=@('Last Name','First Name','DOB','MRN','Member ID','Payer','Phone','Active Date','Last PCP Visit (TIN)','Last Serial PCP Visit','Last Visit (Any)','Next ACV','Next Appt','Future PCP Visits')
$script:NewListDateColumns=@('DOB','Active Date','Last PCP Visit (TIN)','Last Serial PCP Visit','Last Visit (Any)','Next ACV','Next Appt')
$script:NewListProviderColumns=@('Location','Provider','NPI','Risk Pool')
function Get-NewListPath([string]$Id){if($Id -notmatch '^[a-f0-9]{32}$'){throw 'Invalid list ID.'};return (Join-Path $script:Paths.State ('newlist-'+$Id+'.json'))}
function Get-NewListVisit($p){
 # The visits that count: Export's last QEM visit with any PCP in the assigned TIN, and the latest past Serial Scheduling PCP visit date.
 $d=@(@((Get-P $p 'PcpVisit' $null),(Get-P $p 'SerialLast' $null))|Where-Object{$_});if($d.Count -eq 0){return $null};return ($d|Sort-Object -Descending|Select-Object -First 1)
}
function Test-NewUnseen($p,[DateTime]$Cutoff){if(!(Get-P $p 'NewPatient' $false)){return $false};$v=Get-NewListVisit $p;return (!$v -or ([DateTime]$v).Date -lt $Cutoff.Date)}
function New-NewListRow($p){
 return [ordered]@{'Last Name'=[string]$p.Last;'First Name'=[string]$p.First;'DOB'=[string]$p.DOB;'MRN'=[string]$p.MRN;'Member ID'=[string]$p.MemberID;'Payer'=[string](Get-P $p 'Payer' '');'Phone'=[string](Get-P $p 'Phone' '');'Active Date'=[string](Get-P $p 'ActiveDate' '');'Last PCP Visit (TIN)'=(Get-P $p 'PcpVisit' $null);'Last Serial PCP Visit'=(Get-P $p 'SerialLast' $null);'Last Visit (Any)'=(Get-PatientLastVisit $p);'Next ACV'=(Get-P $p 'NextAcv' $null);'Next Appt'=(Get-P $p 'NextAppt' $null);'Future PCP Visits'=[string](Get-P $p 'FutureVisits' '')}
}
function Get-NewListCell($Row,[string]$Col){$v=$Row[$Col];if($null -eq $v){return ''};if($v -is [DateTime]){return $v.ToString('M/d/yyyy')};return [string]$v}
function New-NewListTotal([object[]]$Rows){$t=[ordered]@{providers=0;newPatients=0;notSeen=0};foreach($r in @($Rows)){if($null -eq $r -or [string](Get-Field $r 'error' '')){continue};$t.providers++;$t.newPatients+=[int](Get-Field $r 'newPatients' 0);$t.notSeen+=[int](Get-Field $r 'notSeen' 0)};return $t}
function Get-NewListOrderedRows($Run){return @(foreach($l in @(Get-P $Run 'locations' @())){if($null -ne $l){@(Get-Field $l 'rows' @()|Where-Object{$null -ne $_})}})}
function Get-NewListRowsFor([hashtable]$Lists,[string]$Npi){if($Lists.ContainsKey($Npi)){return @($Lists[$Npi]|Where-Object{$null -ne $_})};return @()}
function Get-NewListCriteriaText($Run){
 $c=ConvertTo-DateValue (Get-P $Run 'cutoff' $null);$m=[int](Get-P $Run 'months' $script:NewListMonths)
 return ('Patients marked New Patient in Export with no visit on or after '+$(if($c){$c.ToString('M/d/yyyy')}else{'the cutoff'})+' ('+$m+' months before the list was built) in either Export "Last QEM Visit Date with any PCP in assigned TIN" or the Serial Scheduling "PCP Visit Dates". Flagged patients are excluded, as on each dashboard.')
}
function Get-SafeFileName([string]$Name){$n=(($Name -replace '[\\/:*?"<>|\x00-\x1f]',' ') -replace '\s+',' ').Trim().TrimEnd('.').Trim();if(!$n){$n='Provider'};if($n.Length -gt 60){$n=$n.Substring(0,60).Trim()};return $n}
function ConvertTo-NewListSummaryHtml($Run){
 # Providers grouped by location with a subtotal under each and a grand total; shared by the page and the combined files.
 $m=[int](Get-P $Run 'months' $script:NewListMonths);$body=''
 foreach($l in @(Get-P $Run 'locations' @())){
  if($null -eq $l){continue};$ln=[string](Get-Field $l 'location' '');$lt=Get-Field $l 'total' $null;$lp=[int](Get-Field $lt 'providers' 0)
  $body+='<tr class="loc"><th colspan="4">'+(ConvertTo-HtmlEncoded $ln)+'</th></tr>'
  foreach($r in @(Get-Field $l 'rows' @()|Where-Object{$null -ne $_})){$err=[string](Get-Field $r 'error' '');$note=[string](Get-Field $r 'note' '');$body+='<tr><td>'+(ConvertTo-HtmlEncoded (Get-Field $r 'displayName' ''))+$(if($err){'<div class="rowerr">Not listed: '+(ConvertTo-HtmlEncoded $err)+'</div>'}elseif($note){'<div class="rownote">'+(ConvertTo-HtmlEncoded $note)+'</div>'}else{''})+'</td><td class="pool">'+(ConvertTo-HtmlEncoded (Get-Field $r 'riskPool' ''))+'</td><td class="n">'+$(if($err){'-'}else{[int](Get-Field $r 'newPatients' 0)})+'</td><td class="n">'+$(if($err){'-'}else{[int](Get-Field $r 'notSeen' 0)})+'</td></tr>'}
  $body+='<tr class="sub"><td>'+(ConvertTo-HtmlEncoded $ln)+' total</td><td>'+$lp+' provider'+$(if($lp -eq 1){''}else{'s'})+'</td><td class="n">'+[int](Get-Field $lt 'newPatients' 0)+'</td><td class="n">'+[int](Get-Field $lt 'notSeen' 0)+'</td></tr>'
 }
 $t=Get-P $Run 'total' $null;$tn=[int](Get-Field $t 'providers' 0)
 $body+='<tr class="grand"><td>All providers</td><td>'+$tn+' provider'+$(if($tn -eq 1){''}else{'s'})+'</td><td class="n">'+[int](Get-Field $t 'newPatients' 0)+'</td><td class="n">'+[int](Get-Field $t 'notSeen' 0)+'</td></tr>'
 return '<table class="ov np"><thead><tr><th>Provider</th><th>Risk pool</th><th class="n">New patients</th><th class="n">Not seen in '+$m+' months</th></tr></thead><tbody>'+$body+'</tbody></table>'
}
function ConvertTo-NewListPatientsHtml([object[]]$Rows){
 $Rows=@($Rows|Where-Object{$null -ne $_});if($Rows.Count -eq 0){return '<p class="none">No new patients without a qualifying visit in the window.</p>'}
 $h='<table class="pl"><thead><tr><th class="i">#</th>';foreach($c in $script:NewListColumns){$h+='<th>'+(ConvertTo-HtmlEncoded $c)+'</th>'};$h+='</tr></thead><tbody>';$i=0
 foreach($r in $Rows){$i++;$h+='<tr><td class="i">'+$i+'</td>';foreach($c in $script:NewListColumns){$v=Get-NewListCell $r $c;$h+='<td>'+$(if($v){ConvertTo-HtmlEncoded $v}elseif($c -in @('Last PCP Visit (TIN)','Last Visit (Any)')){'<span class="none">None</span>'}else{''})+'</td>'};$h+='</tr>'}
 return ($h+'</tbody></table>')
}
function ConvertTo-NewListSectionHtml($Row,[object[]]$Patients,[bool]$Break,[bool]$Heading){
 $err=[string](Get-Field $Row 'error' '');$note=[string](Get-Field $Row 'note' '');$n=[int](Get-Field $Row 'notSeen' 0);$tot=[int](Get-Field $Row 'newPatients' 0)
 $h='<section class="prov'+$(if($Break){' brk'}else{''})+'">'+$(if($Heading){'<h2>'+(ConvertTo-HtmlEncoded (Get-Field $Row 'displayName' ''))+'</h2>'}else{''})
 $h+='<div class="muted">'+(ConvertTo-HtmlEncoded (Get-Field $Row 'location' ''))+' &middot; '+(ConvertTo-HtmlEncoded (Get-Field $Row 'riskPool' ''))+' &middot; NPI '+(ConvertTo-HtmlEncoded (Get-Field $Row 'npi' ''))+$(if($err){''}else{' &middot; <b>'+$n+' of '+$tot+' new patient'+$(if($tot -eq 1){''}else{'s'})+' not seen</b>'})+'</div>'
 if($err){return ($h+'<p class="rowerr">Not listed: '+(ConvertTo-HtmlEncoded $err)+'</p></section>')}
 if($note){$h+='<p class="rownote">'+(ConvertTo-HtmlEncoded $note)+'</p>'}
 return ($h+(ConvertTo-NewListPatientsHtml $Patients)+'</section>')
}
$script:NewListCss='body{font:11px Segoe UI,Arial;color:#172033;margin:16px}h1{font-size:18px;color:#17365d;margin:0 0 3px}h2{font-size:15px;color:#17365d;margin:18px 0 2px}.muted,.none{color:#667085}p.crit{margin:4px 0 10px}table.pl{border-collapse:collapse;width:100%;margin-top:6px}table.pl th,table.pl td{padding:4px 6px;border-bottom:1px solid #dce4ef;text-align:left;vertical-align:top}table.pl thead th{background:#17365d;color:#fff;font-weight:600;vertical-align:bottom}table.pl .i{color:#667085;text-align:right}table.pl tbody tr:nth-child(even) td{background:#f8fafc}p.rownote,p.rowerr{margin:6px 0}table.ov.np{width:auto;min-width:55%}@page{size:letter landscape;margin:.4in}@media print{body{margin:0;font-size:9px}section.brk{page-break-before:always}table.pl tr,table.ov tr{page-break-inside:avoid}table.pl thead,table.ov thead{display:table-header-group}}'
function ConvertTo-NewListDocument($Run,[object[]]$Rows,[hashtable]$Lists,[string]$Npi=''){
 # The HTML and PDF files: combined (summary first, then each provider on its own printed page) or one provider's list.
 $asOf=ConvertTo-DateValue (Get-P $Run 'asOf' $null);$when=$(if($asOf){$asOf.ToString('g')}else{''});$m=[int](Get-P $Run 'months' $script:NewListMonths)
 $sub='<div class="muted">Generated '+(ConvertTo-HtmlEncoded $when)+' &middot; '+(ConvertTo-HtmlEncoded (Get-OverviewSourcesText $Run))+'</div><p class="crit muted">'+(ConvertTo-HtmlEncoded (Get-NewListCriteriaText $Run))+'</p>'
 if($Npi){
  $row=@($Rows|Where-Object{[string](Get-Field $_ 'npi' '') -eq $Npi})[0]
  $body='<h1>'+(ConvertTo-HtmlEncoded (Get-Field $row 'displayName' ''))+' - new patients not seen in '+$m+' months</h1>'+$sub+(ConvertTo-NewListSectionHtml $row (Get-NewListRowsFor $Lists $Npi) $false $false)
 }else{
  $body='<h1>New patients not seen in '+$m+' months</h1>'+$sub+(ConvertTo-NewListSummaryHtml $Run)
  foreach($r in $Rows){$body+=ConvertTo-NewListSectionHtml $r (Get-NewListRowsFor $Lists ([string](Get-Field $r 'npi' ''))) $true $true}
 }
 return ('<!doctype html><html><head><meta charset="utf-8"><title>'+(ConvertTo-HtmlEncoded $script:NewListTitle)+'</title><style>'+$script:OverviewCss+$script:NewListCss+'</style></head><body>'+$body+'</body></html>')
}
function ConvertTo-CsvLine([object[]]$Values){return ((@($Values)|ForEach-Object{'"'+([string]$_).Replace('"','""')+'"'}) -join ',')}
function New-NewListCsv([object[]]$Rows,[hashtable]$Lists,[string]$Npi=''){
 # Combined: one flat list with the provider columns first; per provider: that provider's patients only.
 $sb=New-Object Text.StringBuilder;$heads=$(if($Npi){$script:NewListColumns}else{$script:NewListProviderColumns+$script:NewListColumns});[void]$sb.Append((ConvertTo-CsvLine $heads)+"`r`n")
 foreach($r in $Rows){$rn=[string](Get-Field $r 'npi' '');if($Npi -and $rn -ne $Npi){continue};foreach($x in @(Get-NewListRowsFor $Lists $rn)){$vals=@();if(!$Npi){$vals+=@((Get-Field $r 'location' ''),(Get-Field $r 'displayName' ''),$rn,(Get-Field $r 'riskPool' ''))};foreach($c in $script:NewListColumns){$vals+=(Get-NewListCell $x $c)};[void]$sb.Append((ConvertTo-CsvLine $vals)+"`r`n")}}
 return $sb.ToString()
}
function New-NewListWorkbook($Run,[object[]]$Rows,[hashtable]$Lists,[string]$Npi=''){
 # Combined: a Summary sheet (providers by location with subtotals) and one filterable patient sheet; per provider: that provider's sheet.
 try{Add-Type -AssemblyName System.Drawing -ErrorAction Stop}catch{}
 $pkg=New-Object OfficeOpenXml.ExcelPackage;$asOf=ConvertTo-DateValue (Get-P $Run 'asOf' $null);$m=[int](Get-P $Run 'months' $script:NewListMonths)
 $fill={param($Ws,$Row,[int]$LastCol,[string]$Hex,[bool]$Bold)$rng=$Ws.Cells[$Row,1,$Row,$LastCol];$rng.Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid;$h=$Hex.TrimStart('#');$rng.Style.Fill.BackgroundColor.SetColor([Drawing.Color]::FromArgb([Convert]::ToInt32($h.Substring(0,2),16),[Convert]::ToInt32($h.Substring(2,2),16),[Convert]::ToInt32($h.Substring(4,2),16)));if($Bold){$rng.Style.Font.Bold=$true}}
 $head={param($Ws,$Row,[string[]]$Names)for($c=0;$c -lt $Names.Count;$c++){$Ws.Cells[$Row,($c+1)].Value=$Names[$c]};& $fill $Ws $Row $Names.Count '#17365D' $true;$Ws.Cells[$Row,1,$Row,$Names.Count].Style.Font.Color.SetColor([Drawing.Color]::White);$Ws.Cells[$Row,1,$Row,$Names.Count].Style.WrapText=$true}
 $row0=$null;if($Npi){$row0=@($Rows|Where-Object{[string](Get-Field $_ 'npi' '') -eq $Npi})[0]}
 $title=$(if($Npi){[string](Get-Field $row0 'displayName' '')+' - new patients not seen in '+$m+' months'}else{'New patients not seen in '+$m+' months'})
 $intro={param($Ws)$Ws.Cells[1,1].Value=$title;$Ws.Cells[1,1].Style.Font.Bold=$true;$Ws.Cells[1,1].Style.Font.Size=14;$Ws.Cells[2,1].Value=('Generated '+$(if($asOf){$asOf.ToString('g')}else{''})+'. '+(Get-OverviewSourcesText $Run)+'.');$Ws.Cells[3,1].Value=(Get-NewListCriteriaText $Run)}
 if(!$Npi){
  $ws=$pkg.Workbook.Worksheets.Add('Summary');& $intro $ws;$r=5;$sh=@('Location','Provider','NPI','Risk Pool','New Patients','Not Seen','Note');& $head $ws $r $sh;$sumHead=$r
  foreach($l in @(Get-P $Run 'locations' @())){
   if($null -eq $l){continue};$ln=[string](Get-Field $l 'location' '')
   foreach($x in @(Get-Field $l 'rows' @()|Where-Object{$null -ne $_})){$r++;$err=[string](Get-Field $x 'error' '');$ws.Cells[$r,1].Value=$ln;$ws.Cells[$r,2].Value=[string](Get-Field $x 'displayName' '');$ws.Cells[$r,3].Value=[string](Get-Field $x 'npi' '');$ws.Cells[$r,4].Value=[string](Get-Field $x 'riskPool' '');if($err){$ws.Cells[$r,7].Value=('Not listed: '+$err)}else{$ws.Cells[$r,5].Value=[int](Get-Field $x 'newPatients' 0);$ws.Cells[$r,6].Value=[int](Get-Field $x 'notSeen' 0);$note=[string](Get-Field $x 'note' '');if($note){$ws.Cells[$r,7].Value=$note}}}
   $lt=Get-Field $l 'total' $null;$r++;$ws.Cells[$r,1].Value=$ln;$ws.Cells[$r,2].Value=($ln+' total ('+[int](Get-Field $lt 'providers' 0)+')');$ws.Cells[$r,5].Value=[int](Get-Field $lt 'newPatients' 0);$ws.Cells[$r,6].Value=[int](Get-Field $lt 'notSeen' 0);& $fill $ws $r 6 '#F4F7FB' $true
  }
  $t=Get-P $Run 'total' $null;$r++;$ws.Cells[$r,1].Value='All locations';$ws.Cells[$r,2].Value=('All providers ('+[int](Get-Field $t 'providers' 0)+')');$ws.Cells[$r,5].Value=[int](Get-Field $t 'newPatients' 0);$ws.Cells[$r,6].Value=[int](Get-Field $t 'notSeen' 0);& $fill $ws $r 6 '#DFE9F5' $true
  $ws.Column(1).Width=24;$ws.Column(2).Width=30;$ws.Column(3).Width=13;$ws.Column(4).Width=18;$ws.Column(5).Width=12;$ws.Column(6).Width=12;$ws.Column(7).Width=60;$ws.View.FreezePanes(($sumHead+1),3)
 }
 $ws=$pkg.Workbook.Worksheets.Add('New Patients');& $intro $ws
 $lead=@();if(!$Npi){$lead=@($script:NewListProviderColumns)};$heads=@($lead+$script:NewListColumns);$r=5;& $head $ws $r $heads;$headRow=$r;$off=$lead.Count
 foreach($x in $Rows){
  $rn=[string](Get-Field $x 'npi' '');if($Npi -and $rn -ne $Npi){continue}
  foreach($pt in @(Get-NewListRowsFor $Lists $rn)){
   $r++;if(!$Npi){$ws.Cells[$r,1].Value=[string](Get-Field $x 'location' '');$ws.Cells[$r,2].Value=[string](Get-Field $x 'displayName' '');$ws.Cells[$r,3].Value=$rn;$ws.Cells[$r,4].Value=[string](Get-Field $x 'riskPool' '')}
   for($c=0;$c -lt $script:NewListColumns.Count;$c++){$col=$script:NewListColumns[$c];$v=$pt[$col];if($script:NewListDateColumns -contains $col){$d=ConvertTo-DateValue $v;if($d){$ws.Cells[$r,($off+$c+1)].Value=$d.Date;$ws.Cells[$r,($off+$c+1)].Style.Numberformat.Format='m/d/yyyy'};continue};$s=Get-NewListCell $pt $col;if($s){$ws.Cells[$r,($off+$c+1)].Value=$s}}
  }
 }
 $ws.Cells[$headRow,1,[Math]::Max($r,$headRow),$heads.Count].AutoFilter=$true
 for($c=1;$c -le $heads.Count;$c++){$ws.Column($c).Width=14};if(!$Npi){$ws.Column(1).Width=22;$ws.Column(2).Width=28;$ws.Column(4).Width=18};$ws.Column($off+1).Width=18;$ws.Column($off+2).Width=16
 $ws.View.FreezePanes(($headRow+1),($off+3))
 try{return ,$pkg.GetAsByteArray()}finally{$pkg.Dispose()}
}
function Write-NewListFile($Run,[object[]]$Rows,[hashtable]$Lists,[string]$Npi,[string]$Format,[string]$Dest,[string]$StageBase){
 switch($Format){
  'csv'{[IO.File]::WriteAllText($Dest,(New-NewListCsv $Rows $Lists $Npi),(New-Object Text.UTF8Encoding($true)))}
  'xlsx'{[IO.File]::WriteAllBytes($Dest,(New-NewListWorkbook $Run $Rows $Lists $Npi))}
  'html'{[IO.File]::WriteAllText($Dest,(ConvertTo-NewListDocument $Run $Rows $Lists $Npi),(New-Object Text.UTF8Encoding($false)))}
  'pdf'{
   # Rendered under a space-free staging name (the browser command line is not quoted), then moved to its readable name.
   $h=$StageBase+'.html';$pp=$StageBase+'.pdf'
   try{[IO.File]::WriteAllText($h,(ConvertTo-NewListDocument $Run $Rows $Lists $Npi),(New-Object Text.UTF8Encoding($false)));Convert-HtmlPdf $h $pp;if([IO.File]::Exists($Dest)){[IO.File]::Delete($Dest)};[IO.File]::Move($pp,$Dest)}finally{Remove-Item -LiteralPath $h,$pp -Force -ErrorAction SilentlyContinue}
  }
  default{throw ('Unknown format '+$Format)}
 }
}
function Invoke-NewListRun([string]$Id){
 # Runs in its own worker process: reads each source once for every selected provider, finds the unseen new patients, then writes the chosen files.
 $path=Get-NewListPath $Id;$run=Json $path;if(!$run){throw 'List run was not found.'}
 $sw=[Diagnostics.Stopwatch]::StartNew()
 try{
  Set-P $run 'workerPid' $PID;Set-P $run 'startedUtc' ([DateTime]::UtcNow.ToString('o'));Set-RunProgress $run $path 3 'Reading provider profiles'
  $chosen=@(Get-ChosenProfiles @($run.npis))
  $src=Read-GroupSources $chosen $run $path 5 45;$groups=$src.groups
  $asOf=Get-Date;$cutoff=$asOf.Date.AddMonths(-$script:NewListMonths)
  $rowsOut=@();$lists=@{};$i=0
  foreach($p in $chosen){
   $i++;$npi=[string](Get-P $p 'npi' '');$display=[string](Get-P $p 'displayName' $npi);$loc=([string](Get-P $p 'location' '')).Trim();$pool=([string](Get-P $p 'riskPool' '')).Trim()
   Set-RunProgress $run $path ([int](50+15*($i-1)/$chosen.Count)) ('Finding new patients for '+$display+' ('+$i+' of '+$chosen.Count+')')
   $row=[ordered]@{npi=$npi;displayName=$display;location=$(if($loc){$loc}else{'(No location)'});riskPool=$(if($pool){$pool}else{'(No risk pool)'});newPatients=0;notSeen=0;note='';error=''};$lists[$npi]=@()
   try{
    $pf=Get-ProviderFactsFromGroups $groups $p;$act=Get-ActivePatients $pf.facts @(Get-ProviderFlags $npi)
    $new=@($act.active|Where-Object{Get-P $_ 'NewPatient' $false})
    $unseen=@($new|Where-Object{Test-NewUnseen $_ $cutoff}|Sort-Object -Property @{Expression={[string]$_.Last}},@{Expression={[string]$_.First}})
    $row.newPatients=$new.Count;$row.notSeen=$unseen.Count;$row.note=$pf.note;$lists[$npi]=@($unseen|ForEach-Object{New-NewListRow $_})
   }catch{$row.error=$_.Exception.Message;Log 'NEWLIST_PROVIDER' 'FAILED' ($npi+' '+$_.Exception.Message)}
   $rowsOut+=$row
  }
  $locations=@();foreach($ln in @($rowsOut|ForEach-Object{$_.location}|Select-Object -Unique|Sort-Object -Property $script:GroupSort)){$members=@($rowsOut|Where-Object{$_.location -eq $ln}|Sort-Object -Property @{Expression={$_.displayName}});$locations+=[ordered]@{location=$ln;rows=$members;total=(New-NewListTotal $members)}}
  Set-P $run 'locations' $locations;Set-P $run 'total' (New-NewListTotal $rowsOut);Set-P $run 'sources' @($src.sources);Set-P $run 'asOf' ($asOf.ToString('o'));Set-P $run 'cutoff' ($cutoff.ToString('o'));Set-P $run 'months' $script:NewListMonths;Set-P $run 'failedProviders' @($rowsOut|Where-Object{$_.error}).Count
  $ordered=@(Get-NewListOrderedRows $run)
  # Files: <lists folder>\New Patients Not Seen - <stamp>\ holds the combined files and the zip; per-provider files go in its "By provider" folder.
  $stamp=$asOf.ToString('yyyy-MM-dd HHmmss');$folder=Join-Path $script:Paths.NewLists ($script:NewListTitle+' - '+$stamp);$byDir=Join-Path $folder 'By provider';$null=New-Item -ItemType Directory -Path $folder -Force
  $formats=@(@($run.formats)|ForEach-Object{[string]$_});$grouping=[string]$run.grouping;$plan=New-Object Collections.Generic.List[object]
  if($grouping -in @('combined','both')){foreach($fmt in $formats){$plan.Add([ordered]@{kind='combined';npi='';displayName='All providers';format=$fmt;path=(Join-Path $folder ($script:NewListTitle+' - '+$stamp+'.'+$fmt))})}}
  if($grouping -in @('provider','both')){
   $null=New-Item -ItemType Directory -Path $byDir -Force;$used=@{}
   foreach($r in $ordered){if([string]$r.error){continue};$base=Get-SafeFileName ([string]$r.displayName);if($used.ContainsKey($base.ToUpperInvariant())){$base=$base+' ('+$r.npi+')'};$used[$base.ToUpperInvariant()]=$true
    foreach($fmt in $formats){$plan.Add([ordered]@{kind='provider';npi=[string]$r.npi;displayName=[string]$r.displayName;format=$fmt;path=(Join-Path $byDir ($base+' - '+$script:NewListTitle+' - '+$stamp+'.'+$fmt))})}}
  }
  $files=@();$k=0;$pdfError=''
  foreach($f in $plan){
   $k++;Set-RunProgress $run $path ([int](66+30*($k-1)/[Math]::Max(1,$plan.Count))) ('Writing '+$script:NewListFormats[[string]$f.format]+' for '+$f.displayName+' (file '+$k+' of '+$plan.Count+')')
   $entry=[ordered]@{kind=$f.kind;npi=$f.npi;displayName=$f.displayName;format=$f.format;path=$f.path;error=''}
   # After one PDF fails (no Edge/Chrome, or a timeout) the rest are skipped rather than failing one by one.
   if($f.format -eq 'pdf' -and $pdfError){$entry.error=$pdfError}
   else{try{Write-NewListFile $run $ordered $lists $f.npi $f.format $f.path (Join-Path $script:Paths.Staging ('newlist-'+$Id+'-'+$k))}catch{$entry.error=$_.Exception.Message;Remove-Item -LiteralPath $f.path -Force -ErrorAction SilentlyContinue;if($f.format -eq 'pdf'){$pdfError=$entry.error};Log 'NEWLIST_FILE' 'WARN' ([string]$f.path+' '+$entry.error)}}
   $files+=$entry
  }
  $made=@($files|Where-Object{$_.kind -eq 'provider' -and !$_.error}).Count
  if($made -gt 0){
   Set-RunProgress $run $path 97 'Zipping the per-provider files'
   $zip=Join-Path $folder ($script:NewListTitle+' by provider - '+$stamp+'.zip');$entry=[ordered]@{kind='zip';npi='';displayName='All providers';format='zip';path=$zip;error='';count=$made}
   try{try{Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop}catch{};[IO.Compression.ZipFile]::CreateFromDirectory($byDir,$zip)}catch{$entry.error=$_.Exception.Message;Log 'NEWLIST_ZIP' 'WARN' $entry.error}
   $files+=$entry
  }
  Set-P $run 'folder' $folder;Set-P $run 'files' $files;Set-P $run 'pdfError' $pdfError
  $run.state='Completed';$run.percent=100;$run.stage='Lists complete';Set-P $run 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $run 'seconds' ([Math]::Round($sw.Elapsed.TotalSeconds,1));Save-JsonAtomic $path $run
  Log 'NEWLIST_COMPLETED' 'OK' ($chosen.Count.ToString()+' providers, '+$files.Count+' files in '+[Math]::Round($sw.Elapsed.TotalSeconds,1)+'s')
 }catch{$msg=$_.Exception.Message;$cur=Json $path;if(!$cur){$cur=$run};$cur.state='Failed';$cur.stage='Lists failed';Set-P $cur 'errorSummary' $msg;Set-P $cur 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $cur;Log 'NEWLIST_FAILED' 'FAILED' $msg}
}
function Start-NewListWorker([string]$Id){return (Start-RunWorker 'RunNewListId' $Id ('newlist-'+$Id))}
function Get-NewLists{return @(Get-ChildItem $script:Paths.State -Filter 'newlist-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Where-Object{$null -ne $_}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'requestedUtc' '')}} -Descending)}
function Update-NewListLiveness($Run){
 $id=[string]$Run.listId;$reason=Get-RunStallReason $Run $id 'list' 'Build the lists again.' ('newlist-'+$id)
 if(!$reason){return $false}
 $Run.state='Failed';$Run.stage='Lists failed';Set-P $Run 'errorSummary' $reason;Save-JsonAtomic (Get-NewListPath $id) $Run;Log 'NEWLIST_INTERRUPTED' 'WARN' $id;return $true
}
function Get-NewListOutputText($Run){$g=[string](Get-P $Run 'grouping' '');$f=@(@(Get-P $Run 'formats' @())|ForEach-Object{[string]$script:NewListFormats[[string]$_]});return ([string]$script:NewListGroupings[$g]+': '+($f -join ', '))}
function Get-NewListModel([string]$Id,[switch]$Summary){
 $run=Json (Get-NewListPath $Id);if(!$run){throw 'List run was not found.'}
 $null=Update-NewListLiveness $run
 $c=ConvertTo-DateValue (Get-P $run 'cutoff' $null)
 $m=[ordered]@{listId=[string]$run.listId;state=[string]$run.state;percent=[int](Get-P $run 'percent' 0);stage=[string](Get-P $run 'stage' '');requestedUtc=(ConvertTo-IsoText (Get-P $run 'requestedUtc' ''));completedUtc=(ConvertTo-IsoText (Get-P $run 'completedUtc' ''));asOf=(ConvertTo-IsoText (Get-P $run 'asOf' ''));seconds=(Get-P $run 'seconds' $null);npis=@(Get-P $run 'npis' @());providerCount=@(Get-P $run 'npis' @()).Count;selection=(Get-OverviewSelectionText $run);formats=@(@(Get-P $run 'formats' @())|ForEach-Object{[string]$_});grouping=[string](Get-P $run 'grouping' '');outputText=(Get-NewListOutputText $run);errorSummary=[string](Get-P $run 'errorSummary' '');pdfError=[string](Get-P $run 'pdfError' '');failedProviders=[int](Get-P $run 'failedProviders' 0);cutoff=$(if($c){$c.ToString('M/d/yyyy')}else{''});newPatients=$null;notSeen=$null;folder='';combined=@();zip=$null;perProvider=@();tableHtml='';sourcesText='';criteriaText=''}
 if($m.state -eq 'Completed'){
  $t=Get-P $run 'total' $null;$m.newPatients=[int](Get-Field $t 'newPatients' 0);$m.notSeen=[int](Get-Field $t 'notSeen' 0);$m.folder=[string](Get-P $run 'folder' '')
  $files=@(Get-P $run 'files' @());$byNpi=[ordered]@{}
  for($n=0;$n -lt $files.Count;$n++){
   $f=$files[$n];if($null -eq $f){continue};$kind=[string](Get-Field $f 'kind' '');if($Summary -and $kind -eq 'provider'){continue}
   $fp=[string](Get-Field $f 'path' '');$err=[string](Get-Field $f 'error' '');if(!$err -and !(Test-Path -LiteralPath $fp)){$err='It was moved or deleted from the lists folder.'}
   $item=[ordered]@{format=[string](Get-Field $f 'format' '');url=$(if($err){$null}else{'/new-patients.file?id='+$m.listId+'&n='+$n});error=$err;name=[IO.Path]::GetFileName($fp);count=[int](Get-Field $f 'count' 0)}
   if($kind -eq 'combined'){$m.combined+=$item}elseif($kind -eq 'zip'){$m.zip=$item}elseif($kind -eq 'provider'){$k=[string](Get-Field $f 'npi' '');if(!$byNpi.Contains($k)){$byNpi[$k]=[ordered]@{npi=$k;displayName=[string](Get-Field $f 'displayName' '');files=@()}};$byNpi[$k].files+=$item}
  }
  if(!$Summary){$m.perProvider=@($byNpi.Values);$m.tableHtml=ConvertTo-NewListSummaryHtml $run;$m.sourcesText=Get-OverviewSourcesText $run;$m.criteriaText=Get-NewListCriteriaText $run}
 }
 return $m
}
function Get-NewListPageModel{
 $all=@(Get-NewLists);$history=@(foreach($o in $all){if([string](Get-P $o 'state' '') -in @('Completed','Failed')){Get-NewListModel ([string]$o.listId) -Summary}})
 return [ordered]@{providers=@(Get-OverviewProviders);poc=@(Get-PocLatestByLocation);months=$script:NewListMonths;latest=$(if($all.Count){Get-NewListModel ([string]$all[0].listId)}else{$null});history=$history}
}
function New-NewListRun($Body){
 $req=Get-RequestedProviders $Body
 $want=@(@(Get-P $Body 'formats' @())|ForEach-Object{([string]$_).Trim().ToLowerInvariant()});$formats=@(foreach($k in $script:NewListFormats.Keys){if($want -contains $k){$k}})
 if($formats.Count -eq 0){throw 'Choose at least one format.'}
 $grouping=([string](Get-P $Body 'grouping' '')).Trim().ToLowerInvariant();if(!$grouping -or !$script:NewListGroupings.Contains($grouping)){throw 'Choose a combined file, one file per provider, or both.'}
 foreach($o in @(Get-NewLists)){if([string](Get-P $o 'state' '') -in @('Starting','Running') -and !(Update-NewListLiveness $o)){throw 'New patient lists are already being built; wait for them to finish.'}}
 $id=[Guid]::NewGuid().ToString('N');$path=Get-NewListPath $id
 $run=[ordered]@{listVersion=1;listId=$id;state='Starting';percent=1;stage='Launching list worker';requestedUtc=[DateTime]::UtcNow.ToString('o');npis=@($req.npis);selection=$req.selection;formats=$formats;grouping=$grouping;workerPid=0}
 Save-JsonAtomic $path $run;Log 'NEWLIST_REQUESTED' 'OK' (@($req.npis).Count.ToString()+' providers, '+$grouping+', '+($formats -join '+'))
 if($script:NewListInline){Invoke-NewListRun $id}
 else{try{$null=Start-NewListWorker $id}catch{$run.state='Failed';$run.stage='Worker launch failed';$run['errorSummary']=$_.Exception.Message;Save-JsonAtomic $path $run;Log 'NEWLIST_FAILED' 'FAILED' $_.Exception.Message}}
 foreach($old in @(Get-NewLists|Select-Object -Skip $script:NewListKeep)){$oid=[string](Get-P $old 'listId' '');if($oid -match '^[a-f0-9]{32}$'){Remove-Item -LiteralPath (Get-NewListPath $oid) -Force -ErrorAction SilentlyContinue;Remove-Item -LiteralPath (Join-Path $script:Paths.Logs ('newlist-'+$oid+'.out.log')),(Join-Path $script:Paths.Logs ('newlist-'+$oid+'.err.log')) -Force -ErrorAction SilentlyContinue}}
 return (Get-NewListModel $id)
}
function Send-NewListFile($Context,[string]$Id,[string]$N){
 $run=Json (Get-NewListPath $Id);if(!$run -or [string](Get-P $run 'state' '') -ne 'Completed'){throw 'Those lists are not complete.'}
 $i=0;if(![int]::TryParse($N,[ref]$i)){throw 'Invalid file number.'};$files=@(Get-P $run 'files' @());if($i -lt 0 -or $i -ge $files.Count){throw 'Invalid file number.'}
 $f=$files[$i];$fmt=[string](Get-Field $f 'format' '');$types=@{pdf='application/pdf';html='text/html; charset=utf-8';xlsx='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';csv='text/csv; charset=utf-8';zip='application/zip'}
 if(!$types.ContainsKey($fmt)){throw 'Invalid file.'}
 $full=[IO.Path]::GetFullPath([string](Get-Field $f 'path' ''));$root=[IO.Path]::GetFullPath($script:Paths.NewLists).TrimEnd([IO.Path]::DirectorySeparatorChar)+[IO.Path]::DirectorySeparatorChar
 if(!$full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or !(Test-Path -LiteralPath $full)){throw 'That file is not available; it may have been moved or deleted from the lists folder.'}
 $name=[IO.Path]::GetFileName($full) -replace '[^\x20-\x7E]','_' -replace '"',''
 Send-Bytes $Context $types[$fmt] ([IO.File]::ReadAllBytes($full)) ($(if($fmt -in @('pdf','html')){'inline'}else{'attachment'})+'; filename="'+$name+'"')
}
function NewPatientsPage{
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>New Patients</title>
<style>
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px;max-width:1500px}
__TABCSS__
.card{background:white;border:1px solid #dce4ef;border-radius:8px;padding:18px;margin:14px 0}.card h2{margin:0 0 8px;font-size:17px;color:#17365d}.muted{color:#667085}.error{color:#a61b1b}.warn{color:#9a5b00}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:8px 0}
button,a.btn{padding:7px 12px;cursor:pointer;background:#1769aa;color:#fff;border:0;border-radius:5px;font:inherit;text-decoration:none;display:inline-block}button:disabled{opacity:.5;cursor:default}button.alt{background:#e4e9f0;color:#172033}button.go{background:#1d7a3a;font-weight:600;padding:10px 16px}
__RUNCSS__
__OVCSS__
.opts{border-top:1px solid #eef2f7;margin-top:6px;padding-top:4px}.opts b{min-width:64px;color:#17365d}.opts label{cursor:pointer}.row b.lbl{min-width:96px}table.ov.np{width:auto;min-width:55%}
details.pp{margin:6px 0 10px}details.pp summary{cursor:pointer;color:#1769aa}table.pf{border-collapse:collapse;margin-top:6px}table.pf td{padding:3px 10px 3px 0;border-bottom:1px solid #eef2f7}code{background:#eef2f7;padding:1px 5px;border-radius:3px;word-break:break-all}
</style></head><body>
<header><h1>New Patients</h1><p>Draft 5.6 - patients marked New Patient with no PCP visit in the last __MONTHS__ months, listed per provider</p></header>__NAV__
<main>
<section class="card"><h2>Choose providers</h2>
<p class="muted">Tick a risk pool or location to select all of its providers, or tick providers individually. Locations and risk pools come from each saved provider profile.</p>
<div id="picker"><span class="muted">Loading profiles...</span></div>
<div class="opts">
<div class="row"><b>Files</b><label><input type="radio" name="grouping" value="combined"> One combined file</label><label><input type="radio" name="grouping" value="provider"> One file per provider</label><label><input type="radio" name="grouping" value="both"> Both</label></div>
<div class="row"><b>Formats</b><label><input type="checkbox" name="fmt" value="pdf"> PDF</label><label><input type="checkbox" name="fmt" value="html"> HTML</label><label><input type="checkbox" name="fmt" value="xlsx"> Excel (.xlsx)</label><label><input type="checkbox" name="fmt" value="csv"> CSV</label></div>
</div>
<div class="row"><button class="alt" id="all">Select all</button><button class="alt" id="none">Clear</button><button class="go" id="build" disabled>Build lists</button><span id="count" class="muted"></span></div>
<p id="pocLinks" style="display:none"></p>
<p class="muted">A patient is listed when Export marks them New Patient and neither Export's "Last QEM Visit Date with any PCP in assigned TIN" nor a Serial Scheduling PCP visit date falls within the last __MONTHS__ months. Flagged patients are excluded, as on each dashboard. Per-provider files also come together in one zip.</p>
</section>
<section class="card"><h2>Lists</h2><div id="result"><span class="muted">Loading...</span></div></section>
<section class="card"><h2>Previous lists</h2><p class="muted">The last 30 builds, newest first. View shows one again and restores its providers and output choices; its files are kept as built.</p><div id="history"><span class="muted">Loading...</span></div></section>
</main>
<script>
__PICKERJS__
let current=null,timer=null,history=[];
const FMT={pdf:'PDF',html:'HTML',xlsx:'Excel',csv:'CSV',zip:'Zip'};
function busy(){return !!current&&(current.state==='Starting'||current.state==='Running')}
function grouping(){const r=document.querySelector('input[name=grouping]:checked');return r?r.value:''}
function formats(){return [...document.querySelectorAll('input[name=fmt]:checked')].map(c=>c.value)}
function setOutput(g,f){document.querySelectorAll('input[name=grouping]').forEach(r=>{r.checked=r.value===g});document.querySelectorAll('input[name=fmt]').forEach(c=>{c.checked=f.includes(c.value)})}
function changed(){if(!current||current.state!=='Completed')return false;return [...sel].sort().join(',')!==[...current.npis].sort().join(',')||grouping()!==current.grouping||formats().join(',')!==current.formats.join(',')}
function syncChecks(){
 pickerSync();
 const ok=sel.size>0&&formats().length>0&&!!grouping();el('build').disabled=!ok||busy();el('build').textContent=busy()?'Building...':'Build lists';
 const note=el('selNote');if(note)note.style.display=changed()?'':'none';
}
document.querySelector('.opts').addEventListener('change',syncChecks);
function btnLinks(list){return list.map(f=>f.url?'<a class="btn" '+(f.format==='pdf'||f.format==='html'?'target="_blank" ':'')+'href="'+esc(f.url)+'">'+esc(FMT[f.format]||f.format)+'</a>':'<span class="warn" title="'+esc(f.error)+'">'+esc(FMT[f.format]||f.format)+' not available</span>').join(' ')}
function textLinks(list){return list.map(f=>f.url?'<a '+(f.format==='pdf'||f.format==='html'?'target="_blank" ':'')+'href="'+esc(f.url)+'">'+esc(FMT[f.format]||f.format)+'</a>':'<span class="warn" title="'+esc(f.error)+'">'+esc(FMT[f.format]||f.format)+' n/a</span>').join(' &middot; ')}
function plural(n,w){return n+' '+w+(n===1?'':'s')}
function showResult(o){
 current=o;
 if(!o){el('result').innerHTML='<span class="muted">No lists built yet. Choose providers and output above, then select Build lists.</span>';syncChecks();return}
 if(o.state==='Starting'||o.state==='Running'){el('result').innerHTML='<p>Building lists for '+plural(o.providerCount,'provider')+'... '+o.percent+'%</p><div class="progress"><span style="width:'+Math.max(0,Math.min(100,o.percent))+'%"></span></div><span class="muted">'+esc(o.stage)+'</span>';syncChecks();return}
 if(o.state==='Failed'){el('result').innerHTML='<p class="error">The lists could not be built: '+esc(o.errorSummary)+'</p><span class="muted">Requested '+esc(when(o.requestedUtc))+'</span>';syncChecks();return}
 let h='<div class="row"><b>'+plural(o.notSeen,'new patient')+' not seen</b><span class="muted">of '+o.newPatients+' new across '+plural(o.providerCount,'provider')+'; built '+esc(when(o.completedUtc))+(o.seconds!=null?' in '+o.seconds+' s':'')+'</span></div>';
 if(o.combined.length)h+='<div class="row"><b class="lbl">Combined</b>'+btnLinks(o.combined)+'</div>';
 if(o.zip||o.perProvider.length){
  h+='<div class="row"><b class="lbl">Per provider</b>'+(o.zip?(o.zip.url?'<a class="btn" href="'+esc(o.zip.url)+'">Download all ('+plural(o.zip.count,'file')+', zip)</a>':'<span class="warn">Zip not available: '+esc(o.zip.error)+'</span>'):'')+'</div>';
  h+='<details class="pp"><summary>Individual provider files ('+plural(o.perProvider.length,'provider')+')</summary><table class="pf">'+o.perProvider.map(p=>'<tr><td>'+esc(p.displayName)+'</td><td>'+textLinks(p.files)+'</td></tr>').join('')+'</table></details>';
 }
 if(o.pdfError)h+='<p class="warn">PDF files could not be made: '+esc(o.pdfError)+'</p>';
 h+='<p class="muted">Saved in <code>'+esc(o.folder)+'</code></p><p class="muted">'+esc(o.criteriaText)+' '+esc(o.sourcesText)+'.'+(o.failedProviders?' <span class="error">'+plural(o.failedProviders,'provider')+' could not be listed; see the red notes below.</span>':'')+'</p><p class="warn" id="selNote" style="display:none">Your providers or output choices have changed since these lists were built; select Build lists to update them.</p><div class="scroll">'+o.tableHtml+'</div>';
 el('result').innerHTML=h;syncChecks();
}
function historyHtml(list){
 if(!list.length)return '<span class="muted">None yet.</span>';
 return '<table class="hist"><thead><tr><th>Built</th><th>Providers</th><th>Risk pools</th><th>Locations</th><th class="n">Not seen</th><th>Output</th><th></th></tr></thead><tbody>'+list.map(h=>{const showing=current&&current.listId===h.listId;const s=h.selection||{};return '<tr'+(showing?' class="showing"':'')+'><td>'+esc(when(h.completedUtc||h.requestedUtc))+(showing?' <span class="muted">(shown above)</span>':'')+'</td><td>'+h.providerCount+'</td><td>'+esc(s.riskPools||'')+'</td><td>'+esc(s.locations||'')+'</td><td class="n">'+(h.notSeen!=null?h.notSeen:'')+'</td><td>'+esc(h.outputText)+'</td><td>'+(h.state==='Completed'?'<a href="#" data-view="'+esc(h.listId)+'">View</a>'+(h.combined.length?' &middot; '+textLinks(h.combined):'')+(h.zip?' &middot; '+(h.zip.url?'<a href="'+esc(h.zip.url)+'">Zip</a>':'<span class="warn" title="'+esc(h.zip.error)+'">Zip n/a</span>'):''):'<span class="error">Failed: '+esc(h.errorSummary)+'</span>')+'</td></tr>'}).join('')+'</tbody></table>';
}
function renderHistory(){el('history').innerHTML=historyHtml(history)}
async function loadHistory(){try{const m=await api('/api/new-patients');history=m.history;renderHistory()}catch(e){el('history').innerHTML='<p class="error">'+esc(e.message)+'</p>'}}
function restore(o){sel.clear();o.npis.forEach(n=>{if(providers.some(p=>p.npi===n))sel.add(n)});setOutput(o.grouping,o.formats)}
el('history').addEventListener('click',async e=>{const a=e.target.closest('a[data-view]');if(!a)return;e.preventDefault();if(busy())return;try{const o=await api('/api/new-patients-run?id='+encodeURIComponent(a.dataset.view));restore(o);showResult(o);renderHistory();el('result').scrollIntoView({behavior:'smooth'})}catch(err){el('result').innerHTML='<p class="error">'+esc(err.message)+'</p>'}});
let pollErrors=0;
function startPoll(id){if(timer)clearInterval(timer);pollErrors=0;timer=setInterval(async()=>{try{const o=await api('/api/new-patients-run?id='+encodeURIComponent(id));pollErrors=0;showResult(o);if(!busy()){clearInterval(timer);timer=null;loadHistory()}}catch(e){if(++pollErrors<5)return;clearInterval(timer);timer=null;el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>'}},1500)}
el('build').onclick=async()=>{if(el('build').disabled||busy())return;el('build').disabled=true;try{const o=await api('/api/new-patients',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({npis:[...sel],grouping:grouping(),formats:formats()})});showResult(o);if(busy())startPoll(o.listId)}catch(e){el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>';syncChecks()}};
(async()=>{try{const m=await api('/api/new-patients');providers=m.providers;pocReports=m.poc||[];const last=m.latest;if(last)restore(last);else setOutput('combined',['pdf','xlsx']);presetLocation();renderPicker();showResult(last);history=m.history;renderHistory();if(busy())startPoll(last.listId)}catch(e){el('picker').innerHTML='<p class="error">'+esc(e.message)+'</p>'}})();
</script></body></html>
'@
 return $html.Replace('__NAV__',(Get-NavHtml 'newpatients')).Replace('__TABCSS__',$script:TabCss).Replace('__OVCSS__',$script:OverviewCss).Replace('__RUNCSS__',$script:RunPageCss).Replace('__PICKERJS__',$script:PickerJs).Replace('__MONTHS__',[string]$script:NewListMonths)
}
# --- Draft 5.7: POC Coding tab - point-of-care billing and coding behavior for one department of a billing export, each run saved as an HTML report ---
# The billing export is chosen in the browser and kept under poc-coding\source-files (the newest few), its Department column fills the target
# department list, and each department can be associated with a Provider Analysis location (remembered in state\poc-departments.json).
$script:PocKeep=200;$script:PocFileKeep=10;$script:PocInline=$false
$script:PocRequired=@('Service Prov','Service Dt','Procedure','Department','Tx ID','Diagnosis Code','Encounter Num','Work RVU','Procedure Code','Proc Mod')
$script:PocEmPattern='^(9920[2-5]|9921[1-5])$'
$script:PocPrevPattern='^(9938\d|9939\d)$'
$script:PocAwvPattern='^(G0402|G0438|G0439)$'
$script:PocFocusPattern='^(9920[2-5]|9921[1-5]|9938\d|9939\d|G0402|G0438|G0439|G2211|G0447|9940[1-4]|99406|99407|G0446|G0537|G0538|G0444|96127)$'
$script:PocStopWords=@('CRH','PC','PRIMARY','CARE','MEDICAL','MEDICINE','FAMILY','INTERNAL','PEDIATRICS','OF','THE','AT','AND','OFFICE','CLINIC','CENTER','CTR','GROUP','PRACTICE','DEPT','DEPARTMENT','HEALTH')
$script:PocMapPath=Join-Path $script:Paths.State 'poc-departments.json'
function Get-PocPath([string]$Id){if($Id -notmatch '^[a-f0-9]{32}$'){throw 'Invalid analysis ID.'};return (Join-Path $script:Paths.State ('pocrun-'+$Id+'.json'))}
function Get-PocFilePath([string]$Token){if($Token -notmatch '^[a-f0-9]{32}$'){throw 'Invalid file token.'};return (Join-Path $script:Paths.State ('pocfile-'+$Token+'.json'))}
function Get-PocRuns{return @(Get-ChildItem $script:Paths.State -Filter 'pocrun-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Where-Object{$null -ne $_}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'requestedUtc' '')}} -Descending)}
function Get-PocFiles{return @(Get-ChildItem $script:Paths.State -Filter 'pocfile-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Where-Object{$null -ne $_}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'loadedUtc' '')}} -Descending)}
function Get-PocRate([double]$N,[double]$D){if($D -eq 0){return $null};return ($N/$D)}
function Format-PocPct($Value){if($null -eq $Value){return ''};return ([double]$Value*100).ToString('0.0',[Globalization.CultureInfo]::InvariantCulture)+'%'}
function Format-PocCount($Value){return ([double]$Value).ToString('#,0',[Globalization.CultureInfo]::InvariantCulture)}
function Get-ColumnLetter([int]$Column){$s='';while($Column -gt 0){$m=($Column-1)%26;$s=[string][char](65+$m)+$s;$Column=[int][Math]::Floor(($Column-$m)/26)};return $s}
function Get-PocDepartmentLabel([string]$Department){$t=($Department -replace '\s*\[[^\]]*\]\s*$','').Trim();if(!$t){$t=$Department};return $t}
function Find-PocHeader($Worksheet){
 # The row (within the first 25) that carries every required column, with a header -> column map; otherwise the closest row's missing names.
 if($null -eq $Worksheet.Dimension){return $null}
 $maxCol=[Math]::Min(200,$Worksheet.Dimension.End.Column);$limit=[Math]::Min(25,$Worksheet.Dimension.End.Row);$best=$null
 for($row=1;$row -le $limit;$row++){
  $map=@{};for($col=1;$col -le $maxCol;$col++){$t=([string]$Worksheet.Cells[$row,$col].Text).Trim();if($t -and !$map.ContainsKey($t)){$map[$t]=$col}}
  if($map.Count -eq 0){continue}
  $missing=@($script:PocRequired|Where-Object{!$map.ContainsKey($_)})
  if($missing.Count -eq 0){return [ordered]@{row=$row;map=$map;missing=@()}}
  if($null -eq $best -or $missing.Count -lt $best.missing.Count){$best=[ordered]@{row=$row;map=$null;missing=$missing}}
 }
 return $best
}
function Read-PocWorkbook([string]$Path,[switch]$DepartmentsOnly){
 # The first worksheet with every required header. DepartmentsOnly reads just the Department column (for the dropdown); otherwise every
 # required column is read in one bulk range read each, as trimmed text (index 0 = first data row).
 $package=$null;$closest=$null
 try{
  $package=Open-ExcelPackage -Path $Path -ErrorAction Stop
  foreach($ws in $package.Workbook.Worksheets){
   $h=Find-PocHeader $ws;if($null -eq $h){continue}
   if($null -eq $h.map){if($null -eq $closest -or $h.missing.Count -lt $closest.missing.Count){$closest=[ordered]@{sheet=$ws.Name;missing=$h.missing}};continue}
   $first=$h.row+1;$last=$ws.Dimension.End.Row;$cols=@{}
   $want=@($script:PocRequired);if($DepartmentsOnly){$want=@('Department')}
   foreach($name in $want){if($last -ge $first){$cols[$name]=Get-ColumnText $ws $h.map[$name] $first $last}else{$cols[$name]=New-Object string[] 0}}
   return [ordered]@{sheet=$ws.Name;headerRow=$h.row;rowCount=[Math]::Max(0,$last-$h.row);departmentColumn=(Get-ColumnLetter $h.map['Department']);columns=$cols}
  }
 }finally{if($package){Close-ExcelPackage $package -NoSave}}
 if($closest){throw ('This does not look like the billing export: sheet "'+$closest.sheet+'" is missing '+(@($closest.missing|ForEach-Object{'"'+$_+'"'}) -join ', ')+'.')}
 throw ('This does not look like the billing export: no sheet has the columns '+(@($script:PocRequired|ForEach-Object{'"'+$_+'"'}) -join ', ')+'.')
}
function Get-PocDepartmentCounts([string[]]$Values){
 $c=@{};foreach($v in $Values){if(!$v){continue};if($c.ContainsKey($v)){$c[$v]++}else{$c[$v]=1}}
 return @($c.Keys|Sort-Object|ForEach-Object{[ordered]@{name=$_;rows=$c[$_]}})
}
function Get-PocDepartmentMap{
 $m=@{};$j=Json $script:PocMapPath
 if($j){$d=Get-P $j 'departments' $null;if($d){foreach($p in $d.PSObject.Properties){$m[$p.Name]=[string]$p.Value}}}
 return $m
}
function Set-PocDepartmentLocation([string]$Department,[string]$Location){
 # '' is kept too: it records that the department was deliberately left unassociated, so no location is suggested for it again.
 $m=Get-PocDepartmentMap;$m[$Department]=$Location;$o=[ordered]@{};foreach($k in @($m.Keys|Sort-Object)){$o[$k]=$m[$k]}
 Save-JsonAtomic $script:PocMapPath ([ordered]@{pocMapVersion=1;updatedUtc=[DateTime]::UtcNow.ToString('o');departments=$o})
}
function Get-PocLocations{
 # Every location Provider Analysis knows: saved provider profiles (with how many providers each has) and the per-pool location store.
 $names=@{};$count=@{}
 foreach($p in @(Get-OverviewProviders)){$l=[string]$p.location;if($l -eq '(No location)'){continue};if(!$names.ContainsKey($l)){$names[$l]=$l;$count[$l]=0};$count[$l]++}
 $store=Json (Join-Path $script:Paths.State 'locations.json')
 if($store){$pools=Get-P $store 'pools' $null;if($pools){foreach($pp in $pools.PSObject.Properties){foreach($v in @($pp.Value)){$t=([string]$v).Trim();if($t -and !$names.ContainsKey($t)){$names[$t]=$t;$count[$t]=0}}}}}
 return @($names.Keys|Sort-Object|ForEach-Object{[ordered]@{name=$names[$_];providers=$count[$_]}})
}
function Get-PocTokens([string]$Text){return @((($Text -replace '\[[^\]]*\]',' ').ToUpperInvariant() -split '[^A-Z0-9]+')|Where-Object{$_ -and $script:PocStopWords -notcontains $_})}
function Get-PocSuggestedLocation([string]$Department,[object[]]$Locations){
 # Suggested only when every distinctive word of a location's name appears in the department name and no other location matches as well.
 $dt=@(Get-PocTokens $Department);if($dt.Count -eq 0){return ''}
 $best='';$bestN=0;$tie=$false
 foreach($l in @($Locations)){
  $lt=@(Get-PocTokens ([string]$l.name)|Select-Object -Unique);if($lt.Count -eq 0){continue}
  $all=$true;foreach($w in $lt){if($dt -notcontains $w){$all=$false;break}};if(!$all){continue}
  if($lt.Count -gt $bestN){$best=[string]$l.name;$bestN=$lt.Count;$tie=$false}elseif($lt.Count -eq $bestN){$tie=$true}
 }
 if($tie){return ''};return $best
}
function Remove-OldPocFiles{
 # Keeps the newest few loaded exports (and any a running analysis is reading); older ones are deleted with their records.
 $busy=@{};foreach($r in @(Get-PocRuns)){if([string](Get-P $r 'state' '') -in @('Starting','Running')){$busy[[string](Get-P $r 'fileToken' '')]=$true}}
 foreach($f in @(Get-PocFiles|Select-Object -Skip $script:PocFileKeep)){
  $t=[string](Get-P $f 'token' '');if($t -notmatch '^[a-f0-9]{32}$' -or $busy.ContainsKey($t)){continue}
  $p=[string](Get-P $f 'path' '');if($p -and [IO.Path]::GetFullPath($p).StartsWith([IO.Path]::GetFullPath($script:Paths.PocFiles),[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue}
  Remove-Item -LiteralPath (Get-PocFilePath $t) -Force -ErrorAction SilentlyContinue
 }
}
function Receive-PocUpload($Context){
 # Saves the chosen export, checks its columns and lists its departments. A file already loaded (same SHA-256) is reused.
 $name=[Net.WebUtility]::UrlDecode([string]$Context.Request.Headers['X-File-Name']);if([string]::IsNullOrWhiteSpace($name)){throw 'Missing file name.'}
 $name=[IO.Path]::GetFileName($name);if([IO.Path]::GetExtension($name) -ine '.xlsx'){throw 'Only .xlsx files are accepted.'}
 $token=[Guid]::NewGuid().ToString('N');$path=Join-Path $script:Paths.PocFiles ($token+'.xlsx')
 $stream=[IO.File]::Create($path);try{$buffer=New-Object byte[] 1048576;$total=0;while(($read=$Context.Request.InputStream.Read($buffer,0,$buffer.Length)) -gt 0){$total+=$read;if($total -gt 2147483648){throw 'Upload exceeds 2 GB limit.'};$stream.Write($buffer,0,$read)}}finally{$stream.Dispose()}
 try{
  $hash=Get-FileHash256 $path
  foreach($f in @(Get-PocFiles)){
   if([string](Get-P $f 'sha256' '') -ne $hash -or !(Test-Path -LiteralPath ([string](Get-P $f 'path' '')))){continue}
   Remove-Item -LiteralPath $path -Force;Set-P $f 'originalName' $name;Set-P $f 'loadedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic (Get-PocFilePath ([string]$f.token)) $f
   Log 'POC_FILE_LOADED' 'OK' ($name+' (already loaded)');return (Get-PocFileModel ([string]$f.token))
  }
  $wb=Read-PocWorkbook $path -DepartmentsOnly
  $depts=@(Get-PocDepartmentCounts $wb.columns['Department']);if($depts.Count -eq 0){throw ('The Department column (column '+$wb.departmentColumn+') has no values.')}
  $rec=[ordered]@{pocFileVersion=1;token=$token;originalName=$name;path=$path;sha256=$hash;byteLength=(Get-Item -LiteralPath $path).Length;loadedUtc=[DateTime]::UtcNow.ToString('o');sheet=$wb.sheet;headerRow=$wb.headerRow;rowCount=$wb.rowCount;departmentColumn=$wb.departmentColumn;departments=$depts}
  Save-JsonAtomic (Get-PocFilePath $token) $rec
 }catch{Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue;Log 'POC_FILE' 'FAILED' $_.Exception.Message;throw}
 Log 'POC_FILE_LOADED' 'OK' ($name+'; '+$rec.rowCount+' rows, '+$depts.Count+' departments')
 Remove-OldPocFiles
 return (Get-PocFileModel $token)
}
function Get-PocFileModel([string]$Token){
 $f=Json (Get-PocFilePath $Token);if(!$f){throw 'That file is no longer loaded; choose it again.'}
 if(!(Test-Path -LiteralPath ([string](Get-P $f 'path' '')))){throw 'That file was removed from the poc-coding folder; choose it again.'}
 $map=Get-PocDepartmentMap;$locs=@(Get-PocLocations)
 $last='';foreach($r in @(Get-PocRuns)){if([string](Get-P $r 'fileToken' '') -eq $Token){$last=[string](Get-P $r 'department' '');break}}
 $depts=@(foreach($d in @(Get-P $f 'departments' @())){
  $n=[string]$d.name;$mapped=$map.ContainsKey($n);$loc='';if($mapped){$loc=[string]$map[$n]}else{$loc=Get-PocSuggestedLocation $n $locs}
  [ordered]@{name=$n;rows=[int]$d.rows;location=$loc;mapped=$mapped;suggested=(!$mapped -and [bool]$loc)}
 })
 return [ordered]@{token=[string]$f.token;originalName=[string]$f.originalName;loadedUtc=(ConvertTo-IsoText (Get-P $f 'loadedUtc' ''));sheet=[string]$f.sheet;rowCount=[int]$f.rowCount;departmentColumn=[string](Get-P $f 'departmentColumn' 'E');departments=$depts;lastDepartment=$last}
}
function Invoke-PocAnalysis($Columns,[string]$Department,$Run=$null,[string]$Path=''){
 # Ported from POC-Coding-Server.ps1. Per encounter in the department: within each procedure code + modifier, the latest transaction
 # (highest Tx ID) with a non-zero wRVU is the active line, or the latest line of a tracked code when none is non-zero; more than one
 # E/M (or preventive) line keeps only the latest. Diagnosis and procedure text from every row of the encounter feeds the condition proxies.
 $inv=[Globalization.CultureInfo]::InvariantCulture;$ns=[Globalization.NumberStyles]::Any
 $dept=$Columns['Department'];$enc=$Columns['Encounter Num'];$code=$Columns['Procedure Code'];$mod=$Columns['Proc Mod'];$txt=$Columns['Tx ID'];$rvt=$Columns['Work RVU']
 $dx=$Columns['Diagnosis Code'];$proc=$Columns['Procedure'];$prov=$Columns['Service Prov'];$sdt=$Columns['Service Dt'];$n=$dept.Length
 $diag=@{};for($i=0;$i -lt $n;$i++){$e=$enc[$i];if(!$e){continue};$t=' '+$dx[$i]+' '+$proc[$i];if($diag.ContainsKey($e)){[void]$diag[$e].Append($t)}else{$diag[$e]=[Text.StringBuilder]::new($t)}}
 $byEnc=@{};$order=[Collections.Generic.List[string]]::new();$scope=0;$noEnc=0;$tx=New-Object double[] $n;$rvu=New-Object double[] $n;$cd=New-Object string[] $n
 for($i=0;$i -lt $n;$i++){
  if($dept[$i] -ne $Department){continue};$scope++;$e=$enc[$i];if(!$e){$noEnc++;continue}
  $v=0.0;if([double]::TryParse($txt[$i],$ns,$inv,[ref]$v)){$tx[$i]=$v};$v=0.0;if([double]::TryParse($rvt[$i],$ns,$inv,[ref]$v)){$rvu[$i]=$v};$cd[$i]=$code[$i].ToUpperInvariant()
  if(!$byEnc.ContainsKey($e)){$byEnc[$e]=[Collections.Generic.List[int]]::new();$order.Add($e)};$byEnc[$e].Add($i)
 }
 $encs=[Collections.Generic.List[object]]::new();$confEM=0;$confPrev=0;$base=0.0;$k=0
 foreach($e in $order){
  $k++;if($Run -and $Path -and $k%500 -eq 0){Set-RunProgress $Run $Path ([int](55+30*$k/$order.Count)) ('Analyzing encounter '+$k+' of '+$order.Count)}
  $rows=$byEnc[$e];$groups=@{};$gorder=[Collections.Generic.List[string]]::new();$newest=-1
  foreach($i in $rows){$key=$cd[$i]+'|'+$mod[$i].ToUpperInvariant();if(!$groups.ContainsKey($key)){$groups[$key]=[Collections.Generic.List[int]]::new();$gorder.Add($key)};$groups[$key].Add($i);if($newest -lt 0 -or $tx[$i] -gt $tx[$newest]){$newest=$i}}
  $active=[Collections.Generic.List[int]]::new()
  foreach($key in $gorder){
   $latest=-1;$latestNz=-1
   foreach($i in $groups[$key]){if($latest -lt 0 -or $tx[$i] -gt $tx[$latest]){$latest=$i};if($rvu[$i] -ne 0 -and ($latestNz -lt 0 -or $tx[$i] -gt $tx[$latestNz])){$latestNz=$i}}
   if($latestNz -ge 0){$active.Add($latestNz)}elseif($cd[$latest] -match $script:PocFocusPattern){$active.Add($latest)}
  }
  $em=-1;$emN=0;$pv=-1;$pvN=0
  foreach($i in $active){if($cd[$i] -match $script:PocEmPattern){$emN++;if($em -lt 0 -or $tx[$i] -gt $tx[$em]){$em=$i}}elseif($cd[$i] -match $script:PocPrevPattern){$pvN++;if($pv -lt 0 -or $tx[$i] -gt $tx[$pv]){$pv=$i}}}
  if($emN -gt 1){$confEM++};if($pvN -gt 1){$confPrev++}
  $codes=[Collections.Generic.List[string]]::new()
  foreach($i in $active){$c=$cd[$i];if($c -match $script:PocEmPattern -and $i -ne $em){continue};if($c -match $script:PocPrevPattern -and $i -ne $pv){continue};$base+=$rvu[$i];$codes.Add($c)}
  $t=$(if($diag.ContainsKey($e)){$diag[$e].ToString()}else{''})
  $ob=$t -match '(?i)\bE66\b|obesity';$tob=$t -match '(?i)\bF17\b|\bZ72\.0\b|nicotine|tobacco'
  $asc=$t -match '(?i)\bI10\b|\bE78\b|\bE11\b|\bE66\b|\bF17\b|hypertension|hyperlipidemia|diabetes|obesity|nicotine|tobacco';$mh=$t -match '(?i)\bF32\b|\bF33\b|\bF41\b|\bZ13\.31\b|\bZ13\.32\b|depression|anxiety'
  $obc=$false;$tc=$false;$ac=$false;$mc=$false;$hasG=$false;$awv=[Collections.Generic.List[string]]::new()
  foreach($c in $codes){
   if($c -eq 'G2211'){$hasG=$true}
   if($c -eq 'G0447' -or ($ob -and $c -match '^9940[1-4]$')){$obc=$true}
   if($c -eq '99406' -or $c -eq '99407'){$tc=$true}
   if($c -eq 'G0446' -or $c -eq 'G0537' -or $c -eq 'G0538'){$ac=$true}
   if($c -eq 'G0444' -or $c -eq '96127'){$mc=$true}
   if($c -match $script:PocAwvPattern -and !$awv.Contains($c)){$awv.Add($c)}
  }
  $emCode='';$emMod='';$has25=$false;if($em -ge 0){$emCode=$cd[$em];$emMod=$mod[$em];$has25=$emMod -match '(^|[,;\s])25($|[,;\s])'}
  $pvCode='';if($pv -ge 0){$pvCode=$cd[$pv]}
  $provider=(($prov[$newest] -replace '\s*\[[^\]]*\]\s*$','').Trim());if(!$provider){$provider='(No provider)'}
  $date=ConvertTo-DateValue $sdt[$newest]
  $encs.Add([pscustomobject]@{Encounter=$e;Date=$date;Provider=$provider;EM=$emCode;EMMod=$emMod;Has25=$has25;Prev=$pvCode;AWV=($awv -join ', ');G2211=$hasG;P25=[bool]($pvCode -and $emCode -and $has25);P25G=[bool]($pvCode -and $emCode -and $has25 -and $hasG);ObCode=$obc;TobCode=$tc;AscCode=$ac;MHCode=$mc;Ob=$ob;Tob=$tob;Asc=$asc;MH=$mh})
 }
 return [ordered]@{rawRows=$n;scopeRows=$scope;noEncounterRows=$noEnc;encounters=$encs;base=$base;emConflicts=$confEM;prevConflicts=$confPrev}
}
function New-PocSummary($A){
 # Department totals, the G2211 opportunity counts behind the calculator, and the provider scorecard (peer rate = every other provider's E/M encounters).
 $prov=@{};$em=0;$g=0;$pv=0;$p25=0;$p25g=0;$opp1=0;$opp2=0;$from=$null;$to=$null
 foreach($x in $A.encounters){
  if(!$prov.ContainsKey($x.Provider)){$prov[$x.Provider]=[ordered]@{Provider=$x.Provider;Enc=0;EM=0;G=0;Prev=0;P25=0;P25G=0;ObC=0;Ob=0;TobC=0;Tob=0;AscC=0;Asc=0;MHC=0;MH=0}}
  $r=$prov[$x.Provider];$r['Enc']+=1
  if($x.EM){$em++;$r['EM']+=1;if($x.G2211){$g++;$r['G']+=1}}
  if($x.Prev){$pv++;$r['Prev']+=1};if($x.P25){$p25++;$r['P25']+=1};if($x.P25G){$p25g++;$r['P25G']+=1}
  foreach($f in @(@('ObCode','ObC'),@('Ob','Ob'),@('TobCode','TobC'),@('Tob','Tob'),@('AscCode','AscC'),@('Asc','Asc'),@('MHCode','MHC'),@('MH','MH'))){if($x.($f[0])){$r[$f[1]]+=1}}
  if(($x.Prev -or $x.AWV) -and $x.EM -and $x.Has25 -and !$x.G2211){$opp1++}
  if(!$x.Prev -and !$x.AWV -and $x.EM -eq '99214' -and !$x.G2211){$opp2++}
  if($x.Date){if($null -eq $from -or $x.Date -lt $from){$from=$x.Date};if($null -eq $to -or $x.Date -gt $to){$to=$x.Date}}
 }
 $score=@(foreach($r in $prov.Values){
  $otherEm=$em-$r.EM;$rate=Get-PocRate $r.G $r.EM;$peer=Get-PocRate ($g-$r.G) $otherEm
  $r['Rate']=$rate;$r['Peer']=$peer;$r['Gap']=$(if($r.EM -gt 0 -and $otherEm -gt 0){$rate-$peer}else{$null});$r
 })
 $score=@($score|Sort-Object -Property @{Expression={$null -eq $_.Rate}},@{Expression={if($null -eq $_.Rate){0}else{-[double]$_.Rate}}},@{Expression={$_.Provider}})
 return [ordered]@{rawRows=$A.rawRows;scopeRows=$A.scopeRows;noEncounterRows=$A.noEncounterRows;encounters=$A.encounters.Count;providers=$score.Count;em=$em;g2211=$g;rate=(Get-PocRate $g $em);prev=$pv;p25=$p25;p25g=$p25g;wrvu=[Math]::Round($A.base,2);emConflicts=$A.emConflicts;prevConflicts=$A.prevConflicts;opp1=$opp1;opp2=$opp2;dateFrom=$(if($from){$from.ToString('yyyy-MM-dd')}else{''});dateTo=$(if($to){$to.ToString('yyyy-MM-dd')}else{''});scorecard=$score}
}
$script:PocReportCss=':root{--navy:#17365d;--teal:#1f6d7a;--pale:#eaf3f5;--ink:#243447}*{box-sizing:border-box}body{margin:0;background:#f4f7f9;font:14px Segoe UI,Arial,sans-serif;color:var(--ink)}header{background:var(--navy);color:#fff;padding:22px 4vw}header h1{margin:0 0 6px}header .meta{opacity:.9;line-height:1.6}main{padding:20px 4vw 40px}.caveat{background:#fff6da;border-left:5px solid #d7a900;padding:12px;margin:0 0 16px}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.tile,.card{background:#fff;border-radius:8px;padding:16px;box-shadow:0 1px 5px #ccd}.tile b{display:block;font-size:26px;color:var(--teal);font-variant-numeric:tabular-nums}.tile small{color:#667085}.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:16px}.card{margin-top:16px}.grid .card{margin-top:0}h2{color:var(--navy);margin:0 0 10px;font-size:18px}table{width:100%;border-collapse:collapse;background:#fff}th{background:var(--teal);color:#fff;position:sticky;top:0}th,td{padding:7px 8px;border-bottom:1px solid #dde5e8;text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}th:first-child,td:first-child{text-align:left}.scroll{overflow:auto;max-height:560px}.pos{color:#087830}.neg{color:#b42318}.muted{color:#667085}input[type=range]{width:100%}label{display:block;margin:6px 0}.bars{margin-top:6px}.bar{display:grid;grid-template-columns:minmax(120px,230px) 1fr 130px;align-items:center;gap:10px;padding:3px 0}.bn{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.bt{position:relative;height:14px;background:var(--pale)}.bf{position:absolute;left:0;top:0;bottom:0;background:var(--teal);border-radius:0 4px 4px 0;min-width:2px}.bf.zero{background:none}.ref{position:absolute;top:-4px;bottom:-4px;border-left:2px dashed var(--navy)}.bv{text-align:right;font-variant-numeric:tabular-nums}.bv small{color:#667085}.key{color:#667085;margin:4px 0 8px}.key i{display:inline-block;width:0;height:12px;border-left:2px dashed var(--navy);margin:0 6px -1px 0}ul.notes{margin:6px 0 0 18px;padding:0;line-height:1.6}@media(max-width:900px){.kpis,.grid{grid-template-columns:1fr 1fr}.bar{grid-template-columns:110px 1fr 100px}}@media print{th{position:static}.scroll{max-height:none;overflow:visible}}'
function ConvertTo-PocBarsHtml([object[]]$Items,$Reference){
 # One horizontal bar per item (label, share 0..1, value text, hover text); Reference draws the dashed department line on every track.
 $out=''
 foreach($it in $Items){
  $w=[Math]::Max(0,[Math]::Min(100,100*[double]$it.share));$ws=$w.ToString('0.0',[Globalization.CultureInfo]::InvariantCulture)
  $ref='';if($null -ne $Reference){$ref='<span class="ref" style="left:'+(100*[double]$Reference).ToString('0.0',[Globalization.CultureInfo]::InvariantCulture)+'%"></span>'}
  $out+='<div class="bar" title="'+(ConvertTo-HtmlEncoded $it.title)+'"><span class="bn">'+(ConvertTo-HtmlEncoded $it.label)+'</span><span class="bt"><span class="bf'+$(if($w -eq 0){' zero'}else{''})+'" style="width:'+$ws+'%"></span>'+$ref+'</span><span class="bv">'+$it.value+'</span></div>'
 }
 return '<div class="bars">'+$out+'</div>'
}
function ConvertTo-PocReport($Run){
 # The saved report: header, caveat, KPI tiles, preventive adoption, G2211 calculator, charts and the provider scorecard.
 $s=Get-P $Run 'summary' $null;$e={param($v)ConvertTo-HtmlEncoded $v}
 $em=[int]$s.em;$g=[int]$s.g2211;$pv=[int]$s.prev;$p25=[int]$s.p25;$p25g=[int]$s.p25g;$opp1=[int]$s.opp1;$opp2=[int]$s.opp2;$rate=Get-P $s 'rate' $null
 $score=@(Get-P $s 'scorecard' @()|Where-Object{$null -ne $_})
 $loc=[string](Get-P $Run 'location' '');$locProv=@(Get-P $Run 'locationProviders' @())
 $locText=$(if($loc){'Provider Analysis location: <b>'+(& $e $loc)+'</b>'+$(if($locProv.Count){' ('+$locProv.Count+' provider'+$(if($locProv.Count -eq 1){''}else{'s'})+' with profiles)'}else{''})}else{'Not associated with a Provider Analysis location'})
 $range='';if([string]$s.dateFrom){$range='Services '+(ConvertTo-DateText $s.dateFrom)+' - '+(ConvertTo-DateText $s.dateTo)+' &middot; '}
 $built=ConvertTo-DateValue (Get-P $Run 'asOf' $null);if(!$built){$built=Get-Date}
 $meta='Department: <b>'+(& $e $Run.department)+'</b> &middot; '+$locText+'<br>'+$range+'File: '+(& $e $Run.fileName)+' (sheet '+(& $e (Get-P $Run 'sheet' ''))+') &middot; '+(Format-PocCount $s.scopeRows)+' department rows of '+(Format-PocCount $s.rawRows)+' &middot; '+(Format-PocCount $s.encounters)+' encounters &middot; '+$s.providers+' providers &middot; Built '+$built.ToString('M/d/yyyy h:mm tt')
 $tiles=@(@('Eligible encounters',(Format-PocCount $s.encounters),''),@('E/M encounters',(Format-PocCount $em),''),@('G2211 with E/M',(Format-PocCount $g),''),@('Department G2211 rate',(Format-PocPct $rate),'of E/M encounters'),@('Preventive encounters',(Format-PocCount $pv),''),@('Preventive + E/M with 25',(Format-PocCount $p25),(Format-PocPct (Get-PocRate $p25 $pv))+' of preventive'),@('+ G2211 nested',(Format-PocCount $p25g),(Format-PocPct (Get-PocRate $p25g $pv))+' of preventive'),@('Active wRVUs',([double]$s.wrvu).ToString('#,0.0',[Globalization.CultureInfo]::InvariantCulture),''))
 $tilesHtml=($tiles|ForEach-Object{'<div class="tile">'+$_[0]+'<b>'+$_[1]+'</b>'+$(if($_[2]){'<small>'+$_[2]+'</small>'}else{''})+'</div>'}) -join ''
 $funnel=ConvertTo-PocBarsHtml @(
  [ordered]@{label='Any annual preventive';share=$(if($pv){1}else{0});value=(Format-PocCount $pv)+' <small>'+(Format-PocPct (Get-PocRate $pv $pv))+'</small>';title=''+$pv+' preventive encounters'},
  [ordered]@{label='+ E/M with modifier 25';share=$(if($pv){$p25/$pv}else{0});value=(Format-PocCount $p25)+' <small>'+(Format-PocPct (Get-PocRate $p25 $pv))+'</small>';title=''+$p25+' of '+$pv+' preventive encounters also billed an E/M with 25'},
  [ordered]@{label='+ G2211 as well';share=$(if($pv){$p25g/$pv}else{0});value=(Format-PocCount $p25g)+' <small>'+(Format-PocPct (Get-PocRate $p25g $pv))+'</small>';title=''+$p25g+' of '+$pv+' preventive encounters billed preventive + E/M/25 + G2211'}) $null
 $withEm=@($score|Where-Object{[int]$_.EM -gt 0})
 $gBars=$(if($withEm.Count){'<p class="key"><i></i>Department rate '+(Format-PocPct $rate)+'</p>'+(ConvertTo-PocBarsHtml @($withEm|ForEach-Object{[ordered]@{label=[string]$_.Provider;share=[double]$_.Rate;value=(Format-PocPct $_.Rate)+' <small>'+$_.G+' / '+$_.EM+'</small>';title=[string]$_.Provider+': G2211 on '+$_.G+' of '+$_.EM+' E/M encounters ('+(Format-PocPct $_.Rate)+'); peers '+(Format-PocPct $_.Peer)}}) $rate)}else{'<p class="muted">No E/M encounters in this department.</p>'})
 $rows=($score|ForEach-Object{$gap=Get-P $_ 'Gap' $null;$cls=$(if($null -eq $gap){''}elseif([double]$gap -ge 0){'pos'}else{'neg'});'<tr><td>'+(& $e $_.Provider)+'</td><td>'+$_.Enc+'</td><td>'+$_.EM+'</td><td>'+$_.G+'</td><td>'+(Format-PocPct $_.Rate)+'</td><td>'+(Format-PocPct $_.Peer)+'</td><td class="'+$cls+'">'+(Format-PocPct $gap)+'</td><td>'+$_.Prev+'</td><td>'+$_.P25+'</td><td>'+(Format-PocPct (Get-PocRate $_.P25 $_.Prev))+'</td><td>'+$_.P25G+'</td><td>'+(Format-PocPct (Get-PocRate $_.P25G $_.Prev))+'</td><td>'+$_.ObC+' / '+$_.Ob+'</td><td>'+$_.TobC+' / '+$_.Tob+'</td><td>'+$_.AscC+' / '+$_.Asc+'</td><td>'+$_.MHC+' / '+$_.MH+'</td></tr>'}) -join ''
 $notes='<li>Rows without an encounter number: '+(Format-PocCount $s.noEncounterRows)+' (left out).</li><li>Encounters with more than one active E/M line: '+$s.emConflicts+'; with more than one preventive line: '+$s.prevConflicts+'. Only the latest (highest Tx ID) is kept.</li><li>E/M = 99202-99205, 99211-99215. Preventive = 99381-99397. AWV = G0402, G0438, G0439.</li><li>Obesity code = G0447, or 99401-99404 with an obesity diagnosis. Tobacco = 99406/99407. ASCVD = G0446, G0537, G0538. Mental health screening = G0444, 96127. Each proxy is a diagnosis-text match on the encounter (code / proxy columns count encounters).</li><li>Calculator: "Preventive/AWV + E/M/25" = preventive or AWV encounters with an E/M carrying modifier 25 and no G2211; "Non-preventive 99214" = 99214 encounters without preventive, AWV or G2211.</li>'
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>POC Coding - __TITLE__</title><style>__CSS__</style></head><body>
<header><h1>Point-of-Care Coding Behavior</h1><div class="meta">__META__</div></header><main>
<div class="caveat">Diagnosis-text proxies are broad discussion signals only - not confirmed eligibility, medical necessity, denial risk, compliance findings, or proof of incorrect coding. Latest active state uses Tx ID ordering; zero wRVU alone does not erase an earlier non-zero transaction.</div>
<section class="kpis">__TILES__</section>
<section class="grid"><div class="card"><h2>Preventive adoption</h2><table><tr><th>Stage</th><th>Encounters</th><th>% preventive</th></tr><tr><td>Any annual preventive</td><td>__PV__</td><td>__PVPCT__</td></tr><tr><td>Preventive + E/M with 25</td><td>__P25__</td><td>__P25PCT__</td></tr><tr><td>Preventive + E/M/25 + G2211</td><td>__P25G__</td><td>__P25GPCT__</td></tr></table>__FUNNEL__</div>
<div class="card"><h2>G2211 increased-usage calculator</h2><label>Increase dial: <b id="dialOut">100%</b><input id="dial" type="range" min="0" max="100" value="100"></label><label>wRVU per G2211 <input id="wrvu" type="number" step=".01" value=".33" style="width:80px"></label><table><tr><th>Opportunity</th><th>Without G2211</th><th>Added G2211</th><th>Added wRVU</th></tr><tr><td>Preventive/AWV + E/M/25</td><td>__OPP1__</td><td id="a1"></td><td id="w1"></td></tr><tr><td>Non-preventive 99214</td><td>__OPP2__</td><td id="a2"></td><td id="w2"></td></tr><tr><td>Total</td><td>__OPPT__</td><td id="at"></td><td id="wt"></td></tr></table></div></section>
<section class="card"><h2>G2211 with E/M by provider</h2>__GBARS__</section>
<section class="card"><h2>Provider coding scorecard</h2><div class="scroll"><table><thead><tr><th>Provider</th><th>Encounters</th><th>E/M</th><th>G2211</th><th>G2211 rate</th><th>Peer rate</th><th>Gap</th><th>Preventive</th><th>Prev+E/M/25</th><th>Adoption</th><th>+G2211</th><th>Nested rate</th><th>Obesity code/proxy</th><th>Tobacco code/proxy</th><th>ASCVD code/proxy</th><th>Mental health code/proxy</th></tr></thead><tbody>__ROWS__</tbody></table></div></section>
<section class="card"><h2>How this was counted</h2><ul class="notes">__NOTES__</ul></section>
<script>const d=document.getElementById('dial'),w=document.getElementById('wrvu'),O1=__OPP1__,O2=__OPP2__;function put(id,v){document.getElementById(id).textContent=v}function calc(){const p=d.value/100,v=+w.value||0,x1=Math.round(O1*p),x2=Math.round(O2*p);put('dialOut',d.value+'%');put('a1',x1);put('a2',x2);put('at',x1+x2);put('w1',(x1*v).toFixed(1));put('w2',(x2*v).toFixed(1));put('wt',((x1+x2)*v).toFixed(1))}d.oninput=calc;w.oninput=calc;calc()</script>
</main></body></html>
'@
 return $html.Replace('__TITLE__',(& $e (Get-PocDepartmentLabel ([string]$Run.department)))).Replace('__CSS__',$script:PocReportCss).Replace('__META__',$meta).Replace('__TILES__',$tilesHtml).Replace('__PVPCT__',(Format-PocPct (Get-PocRate $pv $pv))).Replace('__P25PCT__',(Format-PocPct (Get-PocRate $p25 $pv))).Replace('__P25GPCT__',(Format-PocPct (Get-PocRate $p25g $pv))).Replace('__PV__',[string]$pv).Replace('__P25G__',[string]$p25g).Replace('__P25__',[string]$p25).Replace('__FUNNEL__',$funnel).Replace('__OPPT__',[string]($opp1+$opp2)).Replace('__OPP1__',[string]$opp1).Replace('__OPP2__',[string]$opp2).Replace('__GBARS__',$gBars).Replace('__ROWS__',$rows).Replace('__NOTES__',$notes)
}
function Invoke-PocRun([string]$Id){
 # Runs in its own worker process: reads the export, analyzes the chosen department and writes the HTML report into the poc-coding folder.
 $path=Get-PocPath $Id;$run=Json $path;if(!$run){throw 'Analysis was not found.'}
 $sw=[Diagnostics.Stopwatch]::StartNew()
 try{
  Set-P $run 'workerPid' $PID;Set-P $run 'startedUtc' ([DateTime]::UtcNow.ToString('o'));Set-RunProgress $run $path 3 'Opening the billing export'
  $f=Json (Get-PocFilePath ([string]$run.fileToken));$src=$(if($f){[string](Get-P $f 'path' '')}else{''})
  if(!$src -or !(Test-Path -LiteralPath $src)){throw 'The billing export is no longer kept; choose the file again.'}
  Set-RunProgress $run $path 8 ('Reading '+(Format-PocCount (Get-P $f 'rowCount' 0))+' rows from '+[string]$run.fileName)
  $wb=Read-PocWorkbook $src
  Set-RunProgress $run $path 55 ('Analyzing '+[string]$run.department)
  $a=Invoke-PocAnalysis $wb.columns ([string]$run.department) $run $path
  if($a.encounters.Count -eq 0){throw ('No encounters were found for '+[string]$run.department+' ('+$a.scopeRows+' rows).')}
  Set-RunProgress $run $path 88 'Summarizing providers'
  Set-P $run 'summary' (New-PocSummary $a);Set-P $run 'sheet' $wb.sheet;Set-P $run 'asOf' ([DateTime]::Now.ToString('o'))
  Set-RunProgress $run $path 94 'Writing the report';$run=Json $path
  $built=ConvertTo-DateValue (Get-P $run 'asOf' $null);if(!$built){$built=Get-Date}
  $base='POC Coding - '+(Get-SafeFileName (Get-PocDepartmentLabel ([string]$run.department)))+' - '+$built.ToString('yyyy-MM-dd HHmmss');$out=Join-Path $script:Paths.Poc ($base+'.html');$n=2
  while(Test-Path -LiteralPath $out){$out=Join-Path $script:Paths.Poc ($base+' ('+$n+').html');$n++}
  [IO.File]::WriteAllText($out,(ConvertTo-PocReport $run),(New-Object Text.UTF8Encoding($false)))
  Set-P $run 'reportPath' $out;$run.state='Completed';$run.percent=100;$run.stage='Analysis complete';Set-P $run 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $run 'seconds' ([Math]::Round($sw.Elapsed.TotalSeconds,1));Save-JsonAtomic $path $run
  Log 'POC_COMPLETED' 'OK' ([string]$run.department+'; '+$run.summary.encounters+' encounters in '+[Math]::Round($sw.Elapsed.TotalSeconds,1)+'s')
 }catch{$msg=$_.Exception.Message;$cur=Json $path;if(!$cur){$cur=$run};$cur.state='Failed';$cur.stage='Analysis failed';Set-P $cur 'errorSummary' $msg;Set-P $cur 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $cur;Log 'POC_FAILED' 'FAILED' $msg}
}
function Start-PocWorker([string]$Id){return (Start-RunWorker 'RunPocId' $Id ('poc-'+$Id))}
function Update-PocLiveness($Run){
 $id=[string]$Run.pocId;$reason=Get-RunStallReason $Run $id 'POC coding' 'Run the analysis again.' ('poc-'+$id)
 if(!$reason){return $false}
 $Run.state='Failed';$Run.stage='Analysis failed';Set-P $Run 'errorSummary' $reason;Save-JsonAtomic (Get-PocPath $id) $Run;Log 'POC_INTERRUPTED' 'WARN' $id;return $true
}
function Get-PocLocationUrls([string]$Location){$q=[Uri]::EscapeDataString($Location);return [ordered]@{overviewUrl='/overview?location='+$q;newPatientsUrl='/new-patients?location='+$q}}
function Get-PocRunModel([string]$Id,[switch]$Summary){
 $run=Json (Get-PocPath $Id);if(!$run){throw 'That analysis was not found.'}
 $null=Update-PocLiveness $run
 $loc=[string](Get-P $run 'location' '');$token=[string](Get-P $run 'fileToken' '')
 $m=[ordered]@{pocId=[string]$run.pocId;state=[string]$run.state;percent=[int](Get-P $run 'percent' 0);stage=[string](Get-P $run 'stage' '');requestedUtc=(ConvertTo-IsoText (Get-P $run 'requestedUtc' ''));completedUtc=(ConvertTo-IsoText (Get-P $run 'completedUtc' ''));seconds=(Get-P $run 'seconds' $null);errorSummary=[string](Get-P $run 'errorSummary' '');fileName=[string](Get-P $run 'fileName' '');fileToken=$token;fileKept=($token -match '^[a-f0-9]{32}$' -and (Test-Path -LiteralPath (Get-PocFilePath $token)));department=[string](Get-P $run 'department' '');location=$loc;locationProviders=@();overviewUrl=$null;newPatientsUrl=$null;summary=$null;reportUrl=$null;reportFile='';reportError='';folder=$script:Paths.Poc}
 if($loc){$u=Get-PocLocationUrls $loc;$m.overviewUrl=$u.overviewUrl;$m.newPatientsUrl=$u.newPatientsUrl;if(!$Summary){$m.locationProviders=@(Get-OverviewProviders|Where-Object{$_.location -eq $loc}|ForEach-Object{[string]$_.displayName}|Sort-Object)}}
 if($m.state -eq 'Completed'){
  $s=Get-P $run 'summary' $null
  if($s){$m.summary=[ordered]@{rawRows=[int]$s.rawRows;scopeRows=[int]$s.scopeRows;encounters=[int]$s.encounters;providers=[int]$s.providers;em=[int]$s.em;g2211=[int]$s.g2211;rate=(Get-P $s 'rate' $null);prev=[int]$s.prev;p25=[int]$s.p25;p25g=[int]$s.p25g;wrvu=[double]$s.wrvu;opp1=[int]$s.opp1;opp2=[int]$s.opp2;dateFrom=[string]$s.dateFrom;dateTo=[string]$s.dateTo}}
  $rp=[string](Get-P $run 'reportPath' '');if($rp -and (Test-Path -LiteralPath $rp)){$m.reportUrl='/poc-coding.report?id='+$m.pocId;$m.reportFile=$rp}else{$m.reportError='The report was moved or deleted from the poc-coding folder.'}
 }
 return $m
}
function Get-PocLatestByLocation{
 # The newest completed report per associated location, for the cross-links on the Overview and New Patients pages.
 $seen=@{};$out=@()
 foreach($r in @(Get-PocRuns)){
  $loc=[string](Get-P $r 'location' '');if(!$loc -or $seen.ContainsKey($loc) -or [string](Get-P $r 'state' '') -ne 'Completed'){continue}
  $rp=[string](Get-P $r 'reportPath' '');if(!$rp -or !(Test-Path -LiteralPath $rp)){continue}
  $seen[$loc]=$true;$out+=[ordered]@{location=$loc;department=[string](Get-P $r 'department' '');completedUtc=(ConvertTo-IsoText (Get-P $r 'completedUtc' ''));url='/poc-coding.report?id='+[string]$r.pocId}
 }
 return @($out)
}
function Get-PocPageModel{
 $runs=@(Get-PocRuns);$history=@(foreach($r in $runs){if([string](Get-P $r 'state' '') -in @('Completed','Failed')){Get-PocRunModel ([string]$r.pocId) -Summary}})
 $latest=$null;if($runs.Count){$latest=Get-PocRunModel ([string]$runs[0].pocId)}
 $files=@(foreach($f in @(Get-PocFiles)){if(Test-Path -LiteralPath ([string](Get-P $f 'path' ''))){[ordered]@{token=[string]$f.token;originalName=[string]$f.originalName;loadedUtc=(ConvertTo-IsoText (Get-P $f 'loadedUtc' ''));rowCount=[int]$f.rowCount;departments=@(Get-P $f 'departments' @()).Count}}})
 # The file shown first: the latest run's export if it is still kept, otherwise the most recently loaded one.
 $current=$null;$want='';if($latest -and $latest.fileKept){$want=$latest.fileToken}elseif($files.Count){$want=$files[0].token}
 if($want){try{$current=Get-PocFileModel $want}catch{$current=$null}}
 return [ordered]@{locations=@(Get-PocLocations);files=$files;currentFile=$current;latest=$latest;history=$history}
}
function New-PocRun($Body){
 $token=[string](Get-P $Body 'token' '');$f=Json (Get-PocFilePath $token);if(!$f -or !(Test-Path -LiteralPath ([string](Get-P $f 'path' '')))){throw 'That file is no longer loaded; choose it again.'}
 $want=([string](Get-P $Body 'department' '')).Trim();if(!$want){throw 'Choose a target department.'}
 $dept='';foreach($d in @(Get-P $f 'departments' @())){if([string]$d.name -eq $want){$dept=[string]$d.name;break}};if(!$dept){throw ('"'+$want+'" is not a department in this file.')}
 $locWant=([string](Get-P $Body 'location' '')).Trim();$loc=''
 if($locWant){foreach($l in @(Get-PocLocations)){if([string]$l.name -eq $locWant){$loc=[string]$l.name;break}};if(!$loc){$map=Get-PocDepartmentMap;if($map.ContainsKey($dept) -and $map[$dept] -eq $locWant){$loc=[string]$map[$dept]}};if(!$loc){throw ('"'+$locWant+'" is not a Provider Analysis location; reload the page.')}}
 foreach($r in @(Get-PocRuns)){if([string](Get-P $r 'state' '') -in @('Starting','Running') -and !(Update-PocLiveness $r)){throw 'An analysis is already running; wait for it to finish.'}}
 Set-PocDepartmentLocation $dept $loc
 $locProv=@();if($loc){$locProv=@(Get-OverviewProviders|Where-Object{$_.location -eq $loc}|ForEach-Object{[string]$_.displayName}|Sort-Object)}
 $id=[Guid]::NewGuid().ToString('N');$path=Get-PocPath $id
 $run=[ordered]@{pocVersion=1;pocId=$id;state='Starting';percent=1;stage='Launching analysis worker';requestedUtc=[DateTime]::UtcNow.ToString('o');fileToken=$token;fileName=[string]$f.originalName;fileSha256=[string](Get-P $f 'sha256' '');department=$dept;location=$loc;locationProviders=$locProv;workerPid=0}
 Save-JsonAtomic $path $run;Log 'POC_REQUESTED' 'OK' ($dept+$(if($loc){' -> '+$loc}else{''}))
 if($script:PocInline){Invoke-PocRun $id}
 else{try{$null=Start-PocWorker $id}catch{$run.state='Failed';$run.stage='Worker launch failed';$run['errorSummary']=$_.Exception.Message;Save-JsonAtomic $path $run;Log 'POC_FAILED' 'FAILED' $_.Exception.Message}}
 # Old run records drop off the list; their HTML reports stay in the poc-coding folder.
 foreach($old in @(Get-PocRuns|Select-Object -Skip $script:PocKeep)){$oid=[string](Get-P $old 'pocId' '');if($oid -match '^[a-f0-9]{32}$'){Remove-Item -LiteralPath (Get-PocPath $oid) -Force -ErrorAction SilentlyContinue;Remove-Item -LiteralPath (Join-Path $script:Paths.Logs ('poc-'+$oid+'.out.log')),(Join-Path $script:Paths.Logs ('poc-'+$oid+'.err.log')) -Force -ErrorAction SilentlyContinue}}
 return (Get-PocRunModel $id)
}
function Send-PocReport($Context,[string]$Id,[bool]$Download){
 $run=Json (Get-PocPath $Id);if(!$run -or [string](Get-P $run 'state' '') -ne 'Completed'){throw 'That analysis is not complete.'}
 $full=[IO.Path]::GetFullPath([string](Get-P $run 'reportPath' ''));$root=[IO.Path]::GetFullPath($script:Paths.Poc).TrimEnd([IO.Path]::DirectorySeparatorChar)+[IO.Path]::DirectorySeparatorChar
 if(!$full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or !(Test-Path -LiteralPath $full)){throw 'The report is not available; it may have been moved or deleted from the poc-coding folder.'}
 $name=[IO.Path]::GetFileName($full) -replace '[^\x20-\x7E]','_' -replace '"',''
 Send-Bytes $Context 'text/html; charset=utf-8' ([IO.File]::ReadAllBytes($full)) ($(if($Download){'attachment'}else{'inline'})+'; filename="'+$name+'"')
}
function PocCodingPage{
 $html=@'
<!doctype html><html><head><meta charset="utf-8"><title>POC Coding</title>
<style>
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}main{padding:22px;max-width:1500px}
__TABCSS__
.card{background:white;border:1px solid #dce4ef;border-radius:8px;padding:18px;margin:14px 0}.card h2{margin:0 0 8px;font-size:17px;color:#17365d}.muted{color:#667085}.error{color:#a61b1b}.warn{color:#9a5b00}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:8px 0}.row b.lbl{min-width:96px;color:#17365d}select{font:inherit;padding:5px;max-width:100%}
button,a.btn{padding:7px 12px;cursor:pointer;background:#1769aa;color:#fff;border:0;border-radius:5px;font:inherit;text-decoration:none;display:inline-block}button:disabled{opacity:.5;cursor:default}button.alt,a.btn.alt{background:#e4e9f0;color:#172033}button.go{background:#1d7a3a;font-weight:600;padding:10px 16px}
__RUNCSS__
.tiles{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:10px;margin:10px 0}.tile{border:1px solid #dce4ef;border-radius:8px;padding:10px 12px}.tile b{display:block;font-size:22px;color:#1f6d7a}.tile small{color:#667085}
.xref{background:#f4f9ff;border:1px solid #cfe0f5;border-radius:8px;padding:10px 12px;margin:10px 0}code{background:#eef2f7;padding:1px 5px;border-radius:3px;word-break:break-all}.sugg{color:#9a5b00}
@media(max-width:800px){.tiles{grid-template-columns:1fr 1fr}}
</style></head><body>
<header><h1>POC Coding</h1><p>Draft 5.7 - point-of-care billing and coding behavior (E/M, G2211, preventive, counseling and screening codes) for one department of a billing export</p></header>__NAV__
<main>
<section class="card"><h2>1. Choose a billing export</h2>
<p class="muted">An .xlsx with the columns Service Prov, Service Dt, Department, Encounter Num, Tx ID, Procedure Code, Proc Mod, Procedure, Diagnosis Code and Work RVU. The file is kept (the newest __FILEKEEP__) so another department can be analyzed without choosing it again.</p>
<div class="row"><button id="pick">Choose file...</button><input id="file" type="file" accept=".xlsx" hidden><span id="recentWrap" style="display:none" class="muted">or a recent file <select id="recent"></select></span></div>
<div id="fileInfo" class="muted">Loading...</div>
</section>
<section class="card"><h2>2. Target department</h2>
<div class="row"><b class="lbl">Department</b><select id="dept" disabled><option value="">Load a file first</option></select></div>
<div class="row"><b class="lbl">Location</b><select id="loc"></select><span id="locNote" class="muted"></span></div>
<p class="muted">Departments come from the file's Department column. The location ties the department to a location in Provider Analysis (from saved provider profiles), so its report links to that location's Overview and New Patients lists. It is remembered for the department and saved with each run. Providers are not matched yet.</p>
<div class="row"><button class="go" id="run" disabled>Analyze department</button></div>
</section>
<section class="card"><h2>Analysis</h2><div id="result"><span class="muted">Loading...</span></div></section>
<section class="card"><h2>Previous runs</h2><p class="muted">Every run, newest first. Each report is saved as an HTML file in the poc-coding folder. View shows a run again and restores its file, department and location.</p><div id="history"><span class="muted">Loading...</span></div></section>
</main>
<script>
const el=id=>document.getElementById(id);
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function api(u,o){const r=await fetch(u,o);const text=await r.text();let j=null;try{j=text?JSON.parse(text):null}catch(e){throw Error('Server returned invalid JSON: '+text.slice(0,200))}if(!r.ok)throw Error((j&&j.error)||('Request failed ('+r.status+')'));return j}
function when(iso){if(!iso)return '';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleString()}
function day(s){if(!s)return '';const p=s.split('-');return p.length===3?(+p[1])+'/'+(+p[2])+'/'+p[0]:s}
function num(n){return n==null?'':Number(n).toLocaleString()}
function pct(v){return v==null?'':(v*100).toFixed(1)+'%'}
function plural(n,w){return num(n)+' '+w+(n===1?'':'s')}
let file=null,files=[],locations=[],current=null,timer=null,history=[];
function busy(){return !!current&&(current.state==='Starting'||current.state==='Running')}
function sync(){el('run').disabled=!file||!el('dept').value||busy();el('run').textContent=busy()?'Analyzing...':'Analyze department'}
function fillLocations(){el('loc').innerHTML='<option value="">(not associated)</option>'+locations.map(l=>'<option value="'+esc(l.name)+'">'+esc(l.name)+(l.providers?' ('+plural(l.providers,'provider')+')':'')+'</option>').join('')}
function fillRecent(){el('recentWrap').style.display=files.length?'':'none';el('recent').innerHTML='<option value="">Choose...</option>'+files.map(f=>'<option value="'+esc(f.token)+'">'+esc(f.originalName)+' - '+esc(when(f.loadedUtc))+'</option>').join('');el('recent').value=file?file.token:''}
function deptInfo(){return file?file.departments.find(d=>d.name===el('dept').value):null}
function setLocation(v){if(v&&![...el('loc').options].some(o=>o.value===v)){el('loc').add(new Option(v+' (no longer in Provider Analysis)',v))}el('loc').value=v||''}
function deptLabel(x){return x.name+' - '+plural(x.rows,'row')+(x.mapped&&x.location?' - '+x.location:'')}
function showFile(f,dept,loc){
 file=f;
 if(!f){el('fileInfo').textContent='No file loaded yet.';el('dept').disabled=true;el('dept').innerHTML='<option value="">Load a file first</option>';el('locNote').textContent='';fillRecent();sync();return}
 el('fileInfo').innerHTML='<b>'+esc(f.originalName)+'</b> &middot; sheet '+esc(f.sheet)+' &middot; '+plural(f.rowCount,'row')+' &middot; '+plural(f.departments.length,'department')+' in column '+esc(f.departmentColumn)+' &middot; loaded '+esc(when(f.loadedUtc));
 const d=el('dept');d.disabled=false;d.innerHTML='<option value="">Choose a department...</option>'+f.departments.map(x=>'<option value="'+esc(x.name)+'">'+esc(deptLabel(x))+'</option>').join('');
 const want=dept||f.lastDepartment||'';d.value=f.departments.some(x=>x.name===want)?want:'';
 deptChanged(loc);fillRecent();
}
function deptChanged(loc){
 const x=deptInfo();
 if(!x){setLocation('');el('locNote').textContent='';sync();return}
 setLocation(loc!==undefined?loc:x.location);
 el('locNote').className=x.suggested&&loc===undefined?'sugg':'muted';
 el('locNote').textContent=loc!==undefined?(x.mapped&&(x.location||'')===(loc||'')?'remembered for this department':'as used in that run'):x.mapped?(x.location?'remembered for this department':'remembered: not associated'):(x.suggested?'suggested from the department name - check it':'no match found; choose one if it applies');
 sync();
}
el('dept').onchange=()=>deptChanged();
el('loc').onchange=()=>{el('locNote').className='muted';el('locNote').textContent='saved with the next run';sync()};
el('pick').onclick=()=>{el('file').value='';el('file').click()};
el('file').onchange=async()=>{const f=el('file').files[0];if(!f)return;el('fileInfo').textContent='Loading '+f.name+' and reading its departments (large files can take a minute)...';el('dept').disabled=true;el('run').disabled=true;try{const j=await api('/api/poc-coding-upload',{method:'POST',headers:{'X-File-Name':encodeURIComponent(f.name)},body:f});await loadPage(false);showFile(j)}catch(e){el('fileInfo').innerHTML='<span class="error">'+esc(e.message)+'</span>';file=null;el('dept').innerHTML='<option value="">Load a file first</option>';fillRecent();sync()}};
el('recent').onchange=async()=>{const t=el('recent').value;if(!t)return;try{showFile(await api('/api/poc-coding-file?token='+encodeURIComponent(t)))}catch(e){el('fileInfo').innerHTML='<span class="error">'+esc(e.message)+'</span>'}};
function tile(label,value,note){return '<div class="tile">'+esc(label)+'<b>'+value+'</b>'+(note?'<small>'+esc(note)+'</small>':'')+'</div>'}
function xrefHtml(o){
 if(!o.location)return '<div class="xref">Not associated with a Provider Analysis location. Choose a location above and analyze again to link this department to its Overview and New Patients lists.</div>';
 const p=o.locationProviders||[];
 return '<div class="xref"><b>Provider Analysis location: '+esc(o.location)+'</b> <span class="muted">'+(p.length?plural(p.length,'provider')+' with profiles: '+esc(p.join(', ')):'no saved provider profiles at this location yet')+'</span><div class="row"><a class="btn alt" id="ovLink" href="'+esc(o.overviewUrl)+'">Overview for '+esc(o.location)+'</a><a class="btn alt" id="npLink" href="'+esc(o.newPatientsUrl)+'">New patient lists for '+esc(o.location)+'</a></div></div>';
}
function showResult(o){
 current=o;
 if(!o){el('result').innerHTML='<span class="muted">No analysis yet. Choose a file and a department above, then select Analyze department.</span>';sync();return}
 if(busy()){el('result').innerHTML='<p>Analyzing '+esc(o.department)+'... '+o.percent+'%</p><div class="progress"><span style="width:'+Math.max(0,Math.min(100,o.percent))+'%"></span></div><span class="muted">'+esc(o.stage)+'</span>';sync();return}
 if(o.state==='Failed'){el('result').innerHTML='<p class="error">The analysis could not be completed: '+esc(o.errorSummary)+'</p><span class="muted">'+esc(o.department)+' &middot; '+esc(o.fileName)+' &middot; requested '+esc(when(o.requestedUtc))+'</span>';sync();return}
 const s=o.summary||{};
 let h='<div class="row"><b>'+esc(o.department)+'</b><span class="muted">built '+esc(when(o.completedUtc))+(o.seconds!=null?' in '+o.seconds+' s':'')+'</span>'+(o.reportUrl?'<a class="btn" id="open" target="_blank" href="'+esc(o.reportUrl)+'">Open report</a><a class="btn alt" href="'+esc(o.reportUrl)+'&download=1">Download HTML</a>':'<span class="warn">'+esc(o.reportError)+'</span>')+'</div>';
 h+='<div class="tiles">'+tile('Eligible encounters',num(s.encounters))+tile('E/M encounters',num(s.em))+tile('G2211 with E/M',num(s.g2211),pct(s.rate)+' of E/M')+tile('Preventive encounters',num(s.prev))+tile('Preventive + E/M with 25',num(s.p25),s.prev?pct(s.p25/s.prev)+' of preventive':'')+tile('+ G2211 nested',num(s.p25g),s.prev?pct(s.p25g/s.prev)+' of preventive':'')+tile('Active wRVUs',s.wrvu==null?'':Number(s.wrvu).toLocaleString(undefined,{minimumFractionDigits:1,maximumFractionDigits:1}))+tile('Providers',num(s.providers))+'</div>';
 h+=xrefHtml(o);
 h+='<p class="muted">'+esc(o.fileName)+(s.dateFrom?' &middot; services '+esc(day(s.dateFrom))+' - '+esc(day(s.dateTo)):'')+' &middot; '+plural(s.scopeRows,'row')+' in this department of '+num(s.rawRows)+(o.reportFile?'<br>Saved as <code>'+esc(o.reportFile)+'</code>':'')+'</p>';
 el('result').innerHTML=h;sync();
}
function historyHtml(list){
 if(!list.length)return '<span class="muted">None yet.</span>';
 return '<table class="hist"><thead><tr><th>Built</th><th>Department</th><th>Location</th><th>File</th><th class="n">Encounters</th><th class="n">G2211 rate</th><th></th></tr></thead><tbody>'+list.map(r=>{const showing=current&&current.pocId===r.pocId;const s=r.summary||{};return '<tr'+(showing?' class="showing"':'')+'><td>'+esc(when(r.completedUtc||r.requestedUtc))+(showing?' <span class="muted">(shown above)</span>':'')+'</td><td>'+esc(r.department)+'</td><td>'+(r.location?'<a href="'+esc(r.overviewUrl)+'" title="Overview for this location">'+esc(r.location)+'</a>':'<span class="muted">-</span>')+'</td><td>'+esc(r.fileName)+'</td><td class="n">'+num(s.encounters)+'</td><td class="n">'+pct(s.rate)+'</td><td>'+(r.state==='Completed'?'<a href="#" data-view="'+esc(r.pocId)+'">View</a>'+(r.reportUrl?' &middot; <a target="_blank" href="'+esc(r.reportUrl)+'">Report</a>':' &middot; <span class="warn">report missing</span>'):'<span class="error">Failed: '+esc(r.errorSummary)+'</span>')+'</td></tr>'}).join('')+'</tbody></table>';
}
function renderHistory(){el('history').innerHTML=historyHtml(history)}
el('history').addEventListener('click',async e=>{const a=e.target.closest('a[data-view]');if(!a)return;e.preventDefault();if(busy())return;try{const o=await api('/api/poc-coding-run?id='+encodeURIComponent(a.dataset.view));if(o.fileKept){try{showFile(await api('/api/poc-coding-file?token='+encodeURIComponent(o.fileToken)),o.department,o.location)}catch(x){}}showResult(o);renderHistory();el('result').scrollIntoView({behavior:'smooth'})}catch(err){el('result').innerHTML='<p class="error">'+esc(err.message)+'</p>'}});
let pollErrors=0;
function startPoll(id){if(timer)clearInterval(timer);pollErrors=0;timer=setInterval(async()=>{try{const o=await api('/api/poc-coding-run?id='+encodeURIComponent(id));pollErrors=0;showResult(o);if(!busy()){clearInterval(timer);timer=null;const keep=el('dept').value,loc=el('loc').value;await loadPage(false);if(file&&o.fileToken===file.token){try{showFile(await api('/api/poc-coding-file?token='+encodeURIComponent(file.token)),keep,loc)}catch(x){}}}}catch(e){if(++pollErrors<5)return;clearInterval(timer);timer=null;el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>'}},1500)}
el('run').onclick=async()=>{if(el('run').disabled||busy())return;el('run').disabled=true;try{const o=await api('/api/poc-coding',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token:file.token,department:el('dept').value,location:el('loc').value})});showResult(o);if(busy())startPoll(o.pocId);else{await loadPage(false)}}catch(e){el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>';sync()}};
async function loadPage(first){
 const m=await api('/api/poc-coding');files=m.files;history=m.history;
 if(first){locations=m.locations;fillLocations();const l=m.latest;showFile(m.currentFile,l&&m.currentFile&&l.fileToken===m.currentFile.token?l.department:'',undefined);showResult(l);if(busy())startPoll(l.pocId)}
 fillRecent();renderHistory();
}
loadPage(true).catch(e=>{el('fileInfo').innerHTML='<p class="error">'+esc(e.message)+'</p>'});
</script></body></html>
'@
 return $html.Replace('__NAV__',(Get-NavHtml 'poc')).Replace('__TABCSS__',$script:TabCss).Replace('__RUNCSS__',$script:RunPageCss).Replace('__FILEKEEP__',[string]$script:PocFileKeep)
}
# --- Draft 6.0: Combine tab - load and fingerprint exports, combine them into one row per patient with an editable, auditable JSON config ---
# A config is a JSON array with one rule object per line (combine\configs\<name>.json), so it survives a round trip through Excel and an
# enterprise LLM: the Rules page copies a JSON request (how the system works, the current config, the files' fingerprints) to the clipboard,
# and the LLM's answer (one rule per row) is pasted back, checked and diffed before it is saved. Every save keeps the previous version.
# Runs happen in a worker and write one workbook: the patient sheet, one sheet per summary view, and audit sheets tracing every value.
$script:CmbFileKeep=40;$script:CmbRunKeep=100;$script:CmbInline=$false;$script:CmbPreviewRows=50
$script:CmbExtensions=@('.xlsx','.csv','.txt','.tsv')
$script:CmbRuleOrder=@('config','source','map','filter','key','field','derive','lookup','sort','output','view','measure')
$script:CmbHows=@('priority','latest','earliest','max','min','sum','avg','count','countDistinct','list','mostCommon','any','all','exists')
$script:CmbAggs=@('count','sum','avg','min','max','countDistinct','percent')
$script:CmbBuiltins=@('_key','_sources','_rows','_matchedBy')
$script:CmbRuleSpec=[ordered]@{
 config=[ordered]@{purpose='Settings for the whole config. Exactly one, normally the first row.';props=[ordered]@{name='required. Name of the config; also its file name.';description='optional. What the config produces.';asOf='optional. Date (YYYY-MM-DD) that today() returns, so a rerun gives the same answer. Default: the day the output is built.';sheet='optional. Name of the patient sheet in the output workbook. Default "Patients".';lineage='optional. true (default) adds an audit sheet showing which source supplied every value; false leaves it out for very large outputs.';booleans='optional. How true/false values are written: "Y/N" (default), "Yes/No", "TRUE/FALSE" or "1/0".'};example='{"rule":"config","name":"Panel reconciliation","description":"One row per attributed patient from the four registry exports","asOf":"2026-09-30","note":"asOf fixed so reruns match the September reports"}'}
 source=[ordered]@{purpose='Says which loaded file is which export. The order of source rules is the default priority: the first is the most trusted.';props=[ordered]@{source='required. Short id used by the other rules: letters, digits and underscore, e.g. "epic".';label='optional. Friendly name shown in the audit.';match='required. Array of column headers the file must contain (case and extra spaces are ignored). Choose headers that tell this export apart from the others.';fileName='optional. File name pattern with * wildcards, e.g. "*Panel*", for when several loaded files have the match columns.';sheet='optional. Worksheet name (.xlsx). Default: the first sheet with the match columns.';headerRow='optional. Row number of the column headers. Default: found automatically.';role='optional. "patients" (default): each row is about a patient. "lookup": a reference table (for example provider to location to region) used with lookup().';key='lookup sources only: the field (from this source''s map rules) that lookup() matches on.';required='optional. true (default) stops the run when no file matches; false skips the source.';unique='optional. Array of fields; rows whose combination of these values appears more than once in this source are left out (and counted), e.g. ["NPI","MemberID"] when a listing must not repeat a patient.';read='optional. "values" (default) reads what the cells hold; "text" (.xlsx) reads what Excel shows (rounded numbers, percent signs, formatted dates), as the provider reports do.'};example='{"rule":"source","source":"epic","label":"Epic attributed panel","match":["MRN","Patient Name","PCP","Last Office Visit"],"note":"Most trusted for demographics and PCP"}'}
 map=[ordered]@{purpose='Copies a source column (or a constant, or an expression) into a named field, cleaning and typing it. A field name used by several sources means the same thing in each, and is combined across them.';props=[ordered]@{source='required. The source id.';field='required. Field name to fill: letters, digits and underscore, starting with a letter.';column='The column header to read. Exactly one of column, columns, value, expr or special is required.';columns='Array of headers whose non-blank values are joined with sep.';sep='Separator for columns. Default one space.';value='A constant for every row, e.g. "Epic" or "2026-09-30".';expr='An expression over this source''s columns and the fields mapped above it.';special='"fileName", "fileDate" (the file''s modified date), "rowNumber" (row number in the file) or "sourceLabel".';clean='optional. Array of cleaning steps applied in order (see cleaningSteps).';type='optional. "text" (default), "number", "date" or "bool". Values that cannot be read as the type become blank and are counted in the audit.'};example='{"rule":"map","source":"epic","field":"MRN","column":"MRN","clean":["digits","stripLeadingZeros"],"note":"Claims file drops leading zeros, so strip them everywhere"}'}
 filter=[ordered]@{purpose='Drops rows of a source (when it names a source; applied after that source''s map rules) or drops patients (no source; applied after the derive rules).';props=[ordered]@{source='optional. With a source: filters that source''s rows. Without: filters patients.';keep='Expression; rows or patients where it is true are kept. Use keep or drop, not both.';drop='Expression; rows or patients where it is true are removed.'};example='{"rule":"filter","source":"epic","drop":"[Patient Status] in (''Deceased'',''Inactive'')","note":"Only active patients"}'}
 key=[ordered]@{purpose='Decides which rows are the same patient. Rows sharing a key value are joined. Key rules are tried in order; a later key never joins two patients that already have different values of an earlier key (that is reported as a key conflict instead). Key values compare ignoring case and extra spaces; clean them in the map rules so the formats agree. Rows with no usable key are left out and counted.';props=[ordered]@{fields='required. Array of fields that together identify a patient, e.g. ["MRN"] or ["LastName","FirstName","DOB"]. All must be non-blank for a row to use this key.';sources='optional. Array of source ids this key applies to. Default: every patient source that maps all the fields.';match='optional. "exact" (default) or "prefix": a row of these sources joins the one row of the "against" sources whose value begins with its own (with several fields, all but the last must be equal). Several such rows leave it out as ambiguous; none leaves it on its own.';against='prefix only: array of source ids to search. Default: the other patient sources that map the fields.'};example='{"rule":"key","fields":["MRN"],"note":"MRN is in every export"}'}
 field=[ordered]@{purpose='Combines one field across all of a patient''s rows. A mapped field with no field rule uses how "priority".';props=[ordered]@{field='required. Name of the combined field.';of='optional. The mapped field to combine. Default: the same name. Lets one mapped field give several combined fields, e.g. ClaimsPCP from PCP with from ["claims"].';how='optional. priority (default), latest, earliest, max, min, sum, avg, count, countDistinct, list, mostCommon, any, all or exists (see howMethods).';from='optional. Array of source ids to use, in priority order. Default: every source that maps the field, in source order.';by='latest and earliest: the date field that orders the rows; it must be mapped in those sources.';sep='list: separator. Default "; ".';conflicts='optional. "report" (default) lists rows whose value differs from the chosen one on the Conflicts audit sheet; "ignore" does not.';type='optional. Converts the combined value: "text", "number", "date" or "bool".'};example='{"rule":"field","field":"PCP","how":"latest","by":"AttributionDate","from":["claims","epic"],"note":"Use whichever source attributed the patient most recently"}'}
 derive=[ordered]@{purpose='Computes a new field for every patient with an expression, after the field rules, in the order listed. It can use any combined field, earlier derived fields and the built-in fields.';props=[ordered]@{field='required. Name of the new field.';expr='required. The expression.';type='optional. Converts the result: "text", "number", "date" or "bool".'};example='{"rule":"derive","field":"AWVDue","expr":"LastAWV is blank or daysBetween(LastAWV, today()) > 365","type":"bool","note":"Due when none in the past year"}'}
 lookup=[ordered]@{purpose='A small translation table written in the config, used with lookup(''table'', value). For a table kept in a file, use a source with role "lookup" instead.';props=[ordered]@{table='required. Table name.';values='required. Object of "input":"output" pairs; inputs match ignoring case.';default='optional. Value when the input is not in the table. Default blank.'};example='{"rule":"lookup","table":"Region","values":{"Middletown":"Hudson Valley","Monroe":"Hudson Valley","Mt Kisco":"Westchester"},"note":"Location to region"}'}
 sort=[ordered]@{purpose='Orders the patient rows. Sort rules apply in the order listed. Default: the order patients were first found.';props=[ordered]@{field='required. Field to sort by.';desc='optional. true sorts largest or latest first.'};example='{"rule":"sort","field":"PCP"}'}
 output=[ordered]@{purpose='One output column of the patient sheet, in the order listed. Without output rules every field is written, in the order defined.';props=[ordered]@{field='required. Field to write (built-in fields allowed).';header='optional. Column title. Default: the field name.';format='optional. Excel number format, e.g. "m/d/yyyy", "0.0", "0%", "@" for text.';width='optional. Column width.'};example='{"rule":"output","field":"LastAWV","header":"Last AWV","format":"m/d/yyyy"}'}
 view=[ordered]@{purpose='A summary sheet that groups patients, e.g. by provider, location, risk pool or region. Its numbers come from measure rules with the same view name.';props=[ordered]@{view='required. View name; also the sheet name (31 characters at most).';groupBy='required. Array of fields to group by, e.g. ["Region","Location"].';where='optional. Expression; only patients where it is true are counted.';total='optional. true (default) adds a total row.';sortBy='optional. A measure name or group field to sort the rows by. Default: the group fields.';desc='optional. true sorts largest first.'};example='{"rule":"view","view":"By provider","groupBy":["Location","PCP"],"sortBy":"Patients","desc":true}'}
 measure=[ordered]@{purpose='One column of a view.';props=[ordered]@{view='required. The view name.';name='required. Column title.';agg='optional. count (default), sum, avg, min, max, countDistinct or percent.';field='sum, avg, min, max and countDistinct: the field to aggregate.';where='optional. Expression; only patients where it is true are counted in this measure.';of='percent: the measure counted on top.';per='percent: the measure divided by.';format='optional. Excel number format, e.g. "#,##0", "0.0", "0.0%".'};example='{"rule":"measure","view":"By provider","name":"AWV due","agg":"count","where":"AWVDue"}'}
}
$script:CmbHowText=[ordered]@{priority='the first non-blank value, trying sources in "from" order, then rows in file order';latest='the value from the row with the newest "by" date (rows without a date come last; ties go to source priority)';earliest='the value from the row with the oldest "by" date';max='the largest value (numbers, dates, or text in alphabetical order)';min='the smallest value';sum='total of the non-blank numbers';avg='average of the non-blank numbers';count='number of rows with a non-blank value';countDistinct='number of different non-blank values';list='the different non-blank values joined with sep, in priority order';mostCommon='the value found in the most rows (ties go to priority)';any='true if any non-blank value is true (Y, Yes, True, 1, X)';all='true if every non-blank value is true';exists='true if the patient has at least one row in any "from" source ("of" is not needed)'}
$script:CmbCleanText=[ordered]@{
 'trim'='remove spaces at both ends (always done first)';'upper'='capitals';'lower'='lower case';'title'='Title Case';'collapseSpaces'='runs of spaces become one space'
 'digits'='keep only 0-9';'alnum'='keep only letters and digits';'stripLeadingZeros'='remove leading zeros';'pad:N'='left-pad with zeros to N characters, e.g. "pad:9"';'left:N'='keep the first N characters'
 'fillDown'='a blank cell takes the value above it (reports that name a group only on its first row)';'removeBracketId'='drop a trailing [id], e.g. "SMITH, JOHN [12345]" becomes "SMITH, JOHN"';'nameKey'='a person''s name as "LAST, FIRST" in capitals (reads "First Last" and "Last, First Middle"; drops MD, DO, NP, PA-C and similar)'
 'blankIf:TEXT'='treat TEXT as blank, e.g. "blankIf:N/A"';'{"replace":"regex","with":"text"}'='regular expression replace';'{"extract":"regex"}'='keep the first group of the match (or the whole match); no match gives blank'
 '{"map":{"from":"to"},"else":"keep"}'='translate values (ignoring case); "else" is "keep" (default), "blank", or a value to use instead'
}
$script:CmbFunctionText=[ordered]@{
 'if(test, a, b)'='a when test is true, otherwise b';'coalesce(a, b, ...)'='the first non-blank value';'isBlank(x)'='true when x is blank'
 'upper(x), lower(x), trim(x), len(x)'='text case, trimming and length';'left(x, n), right(x, n), mid(x, start, n)'='parts of text; start counts from 1'
 'contains(x, t), startsWith(x, t), endsWith(x, t)'='text tests, ignoring case';'matches(x, ''regex'')'='regular expression test, ignoring case';'replace(x, ''regex'', ''with'')'='regular expression replace';'extract(x, ''regex'')'='the first group of the match, or the whole match'
 'number(x), date(x), text(x)'='convert';'round(x, digits), abs(x), min(a, b, ...), max(a, b, ...)'='numbers (min and max also work on dates)'
 'today()'='the config''s asOf date, or the day the output is built';'year(d), month(d), day(d)'='parts of a date';'addDays(d, n), addMonths(d, n), addYears(d, n)'='date arithmetic'
 'daysBetween(a, b), monthsBetween(a, b), yearsBetween(a, b)'='whole days, months or years from a to b; yearsBetween(DOB, today()) is the age'
 'format(x, ''pattern'')'='a date or number as text, e.g. format(LastVisit, ''yyyy-MM'') or format(Score, ''0.0'')'
 'join(sep, a, b, ...)'='the non-blank values joined with sep, e.g. join('', '', A, B)';'latestDate(x, onOrBefore)'='the latest date in a list of dates in text (separated by commas), optionally only dates on or before onOrBefore';'countItems(x, sep, ignore)'='how many non-blank items x has when split by sep (default comma), not counting items equal to ignore'
 'lookup(''table'', x)'='the value for x in a lookup rule''s table';'lookup(''sourceId'', x, ''Field'')'='a field of the row of a lookup source whose key equals x';'lookup(''profile:Export'', name, ''npi'')'='Provider Analysis profiles: the provider whose saved profile uses this name in that source (any source key: Export, PtListQuality, DiabetesScorecard, SerialScheduling, RiskPopulationOutreach, HR-CRH, HR-RIVPHNYCMM); fields npi, provider, location, riskPool';'lookup(''profiles'', npi, ''provider'')'='a saved profile by NPI; fields provider, location, riskPool'
}
$script:CmbGuide=@(
 'A config is a list of rules (JSON objects) that turns several spreadsheet (.xlsx) or text (.csv, .txt) exports into one row per patient, plus optional summary views. The source files are never changed.'
 'Processing order; each step uses only the rules of its kind:'
 'Step 1. source rules say which loaded file is which export, by the column headers it must contain. Their order is the default priority (first = most trusted).'
 'Step 2. map rules copy source columns into named fields, cleaning and typing the values. The same field name in several sources means the same thing and is combined.'
 'Step 3. filter rules that name a source drop rows of that source.'
 'Step 4. key rules decide which rows belong to the same patient (for example the same MRN).'
 'Step 5. field rules combine each field across a patient''s rows (priority, latest, max, list, count and so on). Disagreements are listed on the Conflicts audit sheet. A mapped field without a field rule uses how "priority".'
 'Step 6. derive rules compute new fields per patient with expressions, in the order listed.'
 'Step 7. filter rules without a source drop patients; sort rules order the rows; output rules choose and name the columns of the patient sheet.'
 'Step 8. view and measure rules build summary sheets, for example by provider, location, risk pool or region.'
 'Field names use letters, digits and underscore and start with a letter (MRN, LastVisit, RiskPool). Names are not case-sensitive.'
 'Built-in patient fields: _key (the key value that matched), _sources (ids of the sources the patient appears in), _rows (number of matched rows), _matchedBy (the key rules that joined the rows).'
 'Value types are text (default), number, date and bool. Blank cells stay blank (not zero). Dates are read from Excel dates or text such as 2026-09-30, 9/30/2026, 30-Sep-2026 or 20260930.'
 'Every rule may carry "note": a short reason for the rule, shown beside it in the audit. Always write one.'
 'The output workbook has audit sheets: each rule''s effect (rows dropped, values unreadable, conflicts found), the files used, every conflict, key matches, excluded patients, and which source supplied each value.'
)
$script:CmbExprGuide=@(
 'Expressions are used by map (expr), filter (keep or drop), derive (expr), view (where) and measure (where).'
 'Names: a field such as PCP, or any name in square brackets such as [Last Visit Date]. In a source''s map and filter rules a name can also be one of that source''s column headers.'
 'Text goes in single quotes: ''Inactive''. Write two single quotes for one inside text: ''O''''Brien''. Numbers: 12, 0.5. Also true, false and blank.'
 'Compare with = <> < <= > >= (text ignores case; numbers and dates compare as numbers and dates). Combine tests with and, or, not, and parentheses.'
 'X in (''A'',''B'') is true when X equals any of the values; X not in (...) is the opposite. X is blank and X is not blank test for empty values.'
 'Arithmetic: + - * /. A date plus a number adds days; a date minus a date gives days. & joins text: LastName & '', '' & FirstName.'
 'A comparison with a blank value is false, except that blank <> ''A'' is true; so Status <> ''Inactive'' keeps rows with a blank Status.'
 'A field on its own is a test: AWVDue is true for Y, Yes, True, 1 or X.'
)
$script:CmbAnswerGuide=@(
 'Answer with the complete new config (not just the changes) and nothing else.'
 'Put the answer in a new worksheet, in column A starting at cell A1, one line per row: row 1 is [ then one complete rule object per row, each followed by a comma except the last rule, and the final row is ].'
 'Never split a rule across rows or put two rules in one row. Do not add headings, explanations, blank rows, or quotes around the rows.'
 'Use straight double quotes for JSON, and single quotes for text inside expressions, so nothing needs escaping.'
 'Keep the rules that are still correct, in the same order where possible. Give every rule a short "note" saying why it exists.'
 'Use only the rule types, properties, cleaning steps and functions described here. Column names in match, column and columns must be copied exactly from the headers listed under columns.'
 'If something cannot be decided from this information, still answer with a complete config and put a note beginning QUESTION: on the affected rule saying what needs checking.'
 'If you cannot create a worksheet, reply with the same lines in a single code block.'
)

# Runtime helpers for expressions and the engine. A class keeps per-value calls cheap compared with functions (Windows PowerShell 5.1 compatible).
class CmbRt{
 static [Globalization.CultureInfo]$Inv=[Globalization.CultureInfo]::InvariantCulture
 static [hashtable]$Rx=@{}
 static [string[]]$DateFormats=@('yyyy-MM-dd','yyyy-MM-ddTHH:mm:ss','yyyy-MM-ddTHH:mm:ss.FFFFFFF','yyyy-MM-ddTHH:mm:ss.FFFFFFFK','yyyy-MM-ddTHH:mm:ssK','yyyy-MM-dd HH:mm:ss','yyyy-MM-dd HH:mm','yyyy/M/d','M/d/yyyy','M/d/yyyy H:mm','M/d/yyyy H:mm:ss','M/d/yyyy h:mm tt','M/d/yyyy h:mm:ss tt','M/d/yy','M-d-yyyy','M.d.yyyy','d-MMM-yyyy','d-MMM-yy','d MMM yyyy','MMM d, yyyy','MMMM d, yyyy','MMM d yyyy','yyyyMMdd')
 static [bool] IsBlank([object]$v){if($null -eq $v){return $true};if($v -is [string]){return [string]::IsNullOrWhiteSpace([string]$v)};return $false}
 static [bool] IsNum([object]$v){return ($v -is [double] -or $v -is [int] -or $v -is [long] -or $v -is [decimal] -or $v -is [single])}
 static [object] Num([object]$v){
  if($null -eq $v){return $null};if($v -is [double]){return [double]$v};if([CmbRt]::IsNum($v)){return [double]$v};if($v -is [bool]){if([bool]$v){return [double]1};return [double]0};if($v -is [DateTime]){return $null}
  $s=([string]$v).Trim();if($s.Length -eq 0){return $null}
  $neg=$false;if($s.StartsWith('(') -and $s.EndsWith(')')){$neg=$true;$s=$s.Substring(1,$s.Length-2)}
  $pct=$false;if($s.EndsWith('%')){$pct=$true;$s=$s.Substring(0,$s.Length-1)}
  $s=$s.Replace(',','').Replace('$','').Replace(' ','')
  [double]$d=0;if(-not [double]::TryParse($s,[Globalization.NumberStyles]::Float,[CmbRt]::Inv,[ref]$d)){return $null}
  if($pct){$d=$d/100};if($neg){$d=-$d};return $d
 }
 static [object] Date([object]$v){
  if($null -eq $v){return $null};if($v -is [DateTime]){return [DateTime]$v}
  if([CmbRt]::IsNum($v)){$n=[double]$v;if($n -ge 1000 -and $n -lt 80000){return [DateTime]::FromOADate($n)};return $null}
  $s=([string]$v).Trim();if($s.Length -eq 0){return $null}
  [DateTime]$d=[DateTime]::MinValue
  if([DateTime]::TryParseExact($s,[CmbRt]::DateFormats,[CmbRt]::Inv,[Globalization.DateTimeStyles]::AllowWhiteSpaces,[ref]$d)){return $d}
  [double]$n=0;if($s -match '^\d{4,5}(\.\d+)?$' -and [double]::TryParse($s,[Globalization.NumberStyles]::Float,[CmbRt]::Inv,[ref]$n) -and $n -ge 1000 -and $n -lt 80000){return [DateTime]::FromOADate($n)}
  if($s -match '[A-Za-z]{3}' -and [DateTime]::TryParse($s,[CmbRt]::Inv,[Globalization.DateTimeStyles]::AllowWhiteSpaces,[ref]$d)){return $d}
  return $null
 }
 static [object] Bool([object]$v){
  if($null -eq $v){return $null};if($v -is [bool]){return [bool]$v};if([CmbRt]::IsNum($v)){return ([double]$v -ne 0)}
  $s=([string]$v).Trim().ToUpperInvariant();if($s.Length -eq 0){return $null}
  if($s -eq 'Y' -or $s -eq 'YES' -or $s -eq 'TRUE' -or $s -eq 'T' -or $s -eq '1' -or $s -eq 'X'){return $true}
  if($s -eq 'N' -or $s -eq 'NO' -or $s -eq 'FALSE' -or $s -eq 'F' -or $s -eq '0'){return $false}
  return $null
 }
 static [bool] Truthy([object]$v){
  if($null -eq $v){return $false};if($v -is [bool]){return [bool]$v};if([CmbRt]::IsNum($v)){return ([double]$v -ne 0)};if($v -is [DateTime]){return $true}
  $s=([string]$v).Trim().ToUpperInvariant();return -not ($s.Length -eq 0 -or $s -eq 'N' -or $s -eq 'NO' -or $s -eq 'FALSE' -or $s -eq 'F' -or $s -eq '0')
 }
 static [string] Text([object]$v){
  if($null -eq $v){return ''};if($v -is [string]){return [string]$v}
  if($v -is [DateTime]){$d=[DateTime]$v;if($d.TimeOfDay.Ticks -eq 0){return $d.ToString('M/d/yyyy',[CmbRt]::Inv)};return $d.ToString('M/d/yyyy H:mm',[CmbRt]::Inv)}
  if([CmbRt]::IsNum($v)){return ([double]$v).ToString('0.###############',[CmbRt]::Inv)};if($v -is [bool]){if([bool]$v){return 'Y'};return 'N'}
  return [string]$v
 }
 static [string] Key([object]$v){$s=[CmbRt]::Text($v).Trim();if($s.Length -eq 0){return ''};return ([regex]::Replace($s,'\s+',' ')).ToUpperInvariant()}
 static [object] Cmp([object]$a,[object]$b){
  # -1/0/1, or $null when either side is blank. Same-type values take the fast paths.
  if($null -eq $a -or $null -eq $b){return $null}
  if($a -is [double] -and $b -is [double]){return ([double]$a).CompareTo([double]$b)}
  if($a -is [DateTime] -and $b -is [DateTime]){return ([DateTime]$a).CompareTo([DateTime]$b)}
  if($a -is [string] -and $b -is [string]){$x=([string]$a).Trim();$y=([string]$b).Trim();if($x.Length -eq 0 -or $y.Length -eq 0){return $null};$r=[string]::Compare($x,$y,[StringComparison]::OrdinalIgnoreCase);if($r -lt 0){return -1};if($r -gt 0){return 1};return 0}
  if([CmbRt]::IsBlank($a) -or [CmbRt]::IsBlank($b)){return $null}
  if($a -is [bool] -or $b -is [bool]){return ([int][CmbRt]::Truthy($a)).CompareTo([int][CmbRt]::Truthy($b))}
  if($a -is [DateTime] -or $b -is [DateTime]){$x=[CmbRt]::Date($a);$y=[CmbRt]::Date($b);if($null -ne $x -and $null -ne $y){return ([DateTime]$x).CompareTo([DateTime]$y)}}
  elseif([CmbRt]::IsNum($a) -or [CmbRt]::IsNum($b)){$x=[CmbRt]::Num($a);$y=[CmbRt]::Num($b);if($null -ne $x -and $null -ne $y){return ([double]$x).CompareTo([double]$y)}}
  $r=[string]::Compare([CmbRt]::Text($a).Trim(),[CmbRt]::Text($b).Trim(),[StringComparison]::OrdinalIgnoreCase);if($r -lt 0){return -1};if($r -gt 0){return 1};return 0
 }
 static [bool] Eq([object]$a,[object]$b){
  if($a -is [string] -and $b -is [string]){$x=([string]$a).Trim();$y=([string]$b).Trim();return [string]::Equals($x,$y,[StringComparison]::OrdinalIgnoreCase)}
  if($a -is [double] -and $b -is [double]){return ([double]$a -eq [double]$b)}
  $ba=[CmbRt]::IsBlank($a);$bb=[CmbRt]::IsBlank($b);if($ba -or $bb){return ($ba -and $bb)};return ([int][CmbRt]::Cmp($a,$b) -eq 0)
 }
 static [bool] Lt([object]$a,[object]$b){$c=[CmbRt]::Cmp($a,$b);return ($null -ne $c -and [int]$c -lt 0)}
 static [bool] Le([object]$a,[object]$b){$c=[CmbRt]::Cmp($a,$b);return ($null -ne $c -and [int]$c -le 0)}
 static [bool] Gt([object]$a,[object]$b){$c=[CmbRt]::Cmp($a,$b);return ($null -ne $c -and [int]$c -gt 0)}
 static [bool] Ge([object]$a,[object]$b){$c=[CmbRt]::Cmp($a,$b);return ($null -ne $c -and [int]$c -ge 0)}
 static [bool] In([object]$a,[object[]]$list){foreach($x in $list){if([CmbRt]::Eq($a,$x)){return $true}};return $false}
 static [string] Cat([object]$a,[object]$b){return [CmbRt]::Text($a)+[CmbRt]::Text($b)}
 static [object] Add([object]$a,[object]$b){
  if([CmbRt]::IsBlank($a) -or [CmbRt]::IsBlank($b)){return $null}
  if($a -is [DateTime]){$n=[CmbRt]::Num($b);if($null -eq $n){return $null};return ([DateTime]$a).AddDays([double]$n)}
  if($b -is [DateTime]){$n=[CmbRt]::Num($a);if($null -eq $n){return $null};return ([DateTime]$b).AddDays([double]$n)}
  $x=[CmbRt]::Num($a);$y=[CmbRt]::Num($b);if($null -ne $x -and $null -ne $y){return ([double]$x+[double]$y)};return [CmbRt]::Text($a)+[CmbRt]::Text($b)
 }
 static [object] Sub([object]$a,[object]$b){
  if([CmbRt]::IsBlank($a) -or [CmbRt]::IsBlank($b)){return $null}
  if($a -is [DateTime]){if($b -is [DateTime]){return [Math]::Round((([DateTime]$a)-([DateTime]$b)).TotalDays,6)};$d=[CmbRt]::Date($b);if($null -ne $d -and -not [CmbRt]::IsNum($b)){return [Math]::Round((([DateTime]$a)-([DateTime]$d)).TotalDays,6)};$n=[CmbRt]::Num($b);if($null -eq $n){return $null};return ([DateTime]$a).AddDays(-[double]$n)}
  $x=[CmbRt]::Num($a);$y=[CmbRt]::Num($b);if($null -eq $x -or $null -eq $y){return $null};return ([double]$x-[double]$y)
 }
 static [object] Mul([object]$a,[object]$b){$x=[CmbRt]::Num($a);$y=[CmbRt]::Num($b);if($null -eq $x -or $null -eq $y){return $null};return ([double]$x*[double]$y)}
 static [object] Div([object]$a,[object]$b){$x=[CmbRt]::Num($a);$y=[CmbRt]::Num($b);if($null -eq $x -or $null -eq $y -or [double]$y -eq 0){return $null};return ([double]$x/[double]$y)}
 static [object] Neg([object]$a){$x=[CmbRt]::Num($a);if($null -eq $x){return $null};return -[double]$x}
 static [object] JoinList([object]$sep,[object[]]$list){$p=[Collections.Generic.List[string]]::new();foreach($x in $list){if(-not [CmbRt]::IsBlank($x)){$p.Add([CmbRt]::Text($x).Trim())}};if($p.Count -eq 0){return $null};return [string]::Join([CmbRt]::Text($sep),$p)}
 static [object] LatestDate([object]$text,[object]$upTo){
  if([CmbRt]::IsBlank($text)){return $null};$lim=$null;if(-not [CmbRt]::IsBlank($upTo)){$lim=[CmbRt]::Date($upTo)};$best=$null
  foreach($part in [CmbRt]::Text($text).Split([char[]]@([char]',',[char]';'))){$d=[CmbRt]::Date($part.Trim());if($null -eq $d){continue};if($null -ne $lim -and [DateTime]$d -gt [DateTime]$lim){continue};if($null -eq $best -or [DateTime]$d -gt [DateTime]$best){$best=$d}}
  return $best
 }
 static [object] CountItems([object]$text,[object]$sep,[object]$ignore){
  if([CmbRt]::IsBlank($text)){return [double]0};$sp=[CmbRt]::Text($sep);if($sp.Length -eq 0){$sp=','};$ig=[CmbRt]::Text($ignore).Trim();$n=0
  foreach($part in [CmbRt]::Text($text).Split([string[]]@($sp),[StringSplitOptions]::None)){$t=$part.Trim();if($t.Length -gt 0 -and ($ig.Length -eq 0 -or $t -ne $ig)){$n++}};return [double]$n
 }
 static [object] Coalesce([object[]]$list){foreach($x in $list){if(-not [CmbRt]::IsBlank($x)){return $x}};return $null}
 static [object] Pick([object[]]$list,[int]$sign){
  $best=$null;foreach($x in $list){if([CmbRt]::IsBlank($x)){continue};if($null -eq $best){$best=$x;continue};$c=[CmbRt]::Cmp($x,$best);if($null -ne $c -and [int]$c*$sign -gt 0){$best=$x}};return $best
 }
 static [object] Min([object[]]$list){return [CmbRt]::Pick($list,-1)}
 static [object] Max([object[]]$list){return [CmbRt]::Pick($list,1)}
 static [object] Blank([string]$s){if($s.Length -eq 0){return $null};return $s}
 static [object] Upper([object]$a){return [CmbRt]::Blank([CmbRt]::Text($a).ToUpperInvariant())}
 static [object] Lower([object]$a){return [CmbRt]::Blank([CmbRt]::Text($a).ToLowerInvariant())}
 static [object] Trim([object]$a){return [CmbRt]::Blank([CmbRt]::Text($a).Trim())}
 static [object] Len([object]$a){return [double][CmbRt]::Text($a).Length}
 static [int] Int([object]$n){$x=[CmbRt]::Num($n);if($null -eq $x){return 0};return [int][Math]::Floor([double]$x)}
 static [object] Left([object]$a,[object]$n){$s=[CmbRt]::Text($a);$k=[Math]::Max(0,[Math]::Min($s.Length,[CmbRt]::Int($n)));return [CmbRt]::Blank($s.Substring(0,$k))}
 static [object] Right([object]$a,[object]$n){$s=[CmbRt]::Text($a);$k=[Math]::Max(0,[Math]::Min($s.Length,[CmbRt]::Int($n)));return [CmbRt]::Blank($s.Substring($s.Length-$k))}
 static [object] Mid([object]$a,[object]$st,[object]$n){$s=[CmbRt]::Text($a);$i=[Math]::Max(1,[CmbRt]::Int($st))-1;if($i -ge $s.Length){return $null};$k=[Math]::Max(0,[Math]::Min($s.Length-$i,[CmbRt]::Int($n)));return [CmbRt]::Blank($s.Substring($i,$k))}
 static [bool] Contains([object]$a,[object]$t){return ([CmbRt]::Text($a).IndexOf([CmbRt]::Text($t),[StringComparison]::OrdinalIgnoreCase) -ge 0)}
 static [bool] StartsWith([object]$a,[object]$t){return [CmbRt]::Text($a).StartsWith([CmbRt]::Text($t),[StringComparison]::OrdinalIgnoreCase)}
 static [bool] EndsWith([object]$a,[object]$t){return [CmbRt]::Text($a).EndsWith([CmbRt]::Text($t),[StringComparison]::OrdinalIgnoreCase)}
 static [regex] Re([string]$p){if(-not [CmbRt]::Rx.ContainsKey($p)){[CmbRt]::Rx[$p]=[regex]::new($p,[Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant')};return [regex][CmbRt]::Rx[$p]}
 static [bool] Matches([object]$a,[object]$p){if([CmbRt]::IsBlank($a)){return $false};return [CmbRt]::Re([CmbRt]::Text($p)).IsMatch([CmbRt]::Text($a))}
 static [object] Replace([object]$a,[object]$p,[object]$w){if([CmbRt]::IsBlank($a)){return $null};return [CmbRt]::Blank([CmbRt]::Re([CmbRt]::Text($p)).Replace([CmbRt]::Text($a),[CmbRt]::Text($w)))}
 static [object] Extract([object]$a,[object]$p){if([CmbRt]::IsBlank($a)){return $null};$m=[CmbRt]::Re([CmbRt]::Text($p)).Match([CmbRt]::Text($a));if(-not $m.Success){return $null};if($m.Groups.Count -gt 1){return [CmbRt]::Blank($m.Groups[1].Value)};return [CmbRt]::Blank($m.Value)}
 static [object] Round([object]$a,[object]$n){$x=[CmbRt]::Num($a);if($null -eq $x){return $null};return [Math]::Round([double]$x,[Math]::Max(0,[Math]::Min(15,[CmbRt]::Int($n))),[MidpointRounding]::AwayFromZero)}
 static [object] Abs([object]$a){$x=[CmbRt]::Num($a);if($null -eq $x){return $null};return [Math]::Abs([double]$x)}
 static [object] Part([object]$d,[string]$what){$x=[CmbRt]::Date($d);if($null -eq $x){return $null};$t=[DateTime]$x;if($what -eq 'y'){return [double]$t.Year};if($what -eq 'm'){return [double]$t.Month};return [double]$t.Day}
 static [object] AddTo([object]$d,[object]$n,[string]$what){
  $x=[CmbRt]::Date($d);$k=[CmbRt]::Num($n);if($null -eq $x -or $null -eq $k){return $null};$t=[DateTime]$x
  if($what -eq 'd'){return $t.AddDays([double]$k)};if($what -eq 'm'){return $t.AddMonths([int][Math]::Truncate([double]$k))};return $t.AddYears([int][Math]::Truncate([double]$k))
 }
 static [object] Between([object]$a,[object]$b,[string]$what){
  if($a -is [DateTime]){$s=([DateTime]$a).Date}else{$x=[CmbRt]::Date($a);if($null -eq $x){return $null};$s=([DateTime]$x).Date}
  if($b -is [DateTime]){$e=([DateTime]$b).Date}else{$y=[CmbRt]::Date($b);if($null -eq $y){return $null};$e=([DateTime]$y).Date}
  if($what -eq 'd'){return [double](($e-$s).Days)}
  $m=($e.Year-$s.Year)*12+$e.Month-$s.Month
  if($e -ge $s){if($e.Day -lt $s.Day){$m--}}elseif($e.Day -gt $s.Day){$m++}
  if($what -eq 'm'){return [double]$m};return [double][Math]::Truncate($m/12)
 }
 static [object] Format([object]$a,[object]$f){
  if([CmbRt]::IsBlank($a)){return $null};$p=[CmbRt]::Text($f)
  if($a -is [DateTime]){return ([DateTime]$a).ToString($p,[CmbRt]::Inv)};$x=[CmbRt]::Num($a);if($null -ne $x){return ([double]$x).ToString($p,[CmbRt]::Inv)}
  $d=[CmbRt]::Date($a);if($null -ne $d){return ([DateTime]$d).ToString($p,[CmbRt]::Inv)};return [CmbRt]::Text($a)
 }
 static [object] Lookup([hashtable]$t,[object]$k,[string]$field){
  $key=[CmbRt]::Key($k);if($key.Length -eq 0 -or -not $t.ContainsKey($key)){if($t.ContainsKey('*default*') -and $field.Length -eq 0){return $t['*default*']};return $null}
  $v=$t[$key];if($field.Length -eq 0){return $v};$h=[hashtable]$v;if($h.ContainsKey($field)){return $h[$field]};return $null
 }
 static [string] NameKey([string]$s){
  # A person's name as "LAST, FIRST" in capitals: drops a trailing [id], punctuation, credentials and suffixes; reads "First M Last" too.
  $t=[regex]::Replace($s.ToUpperInvariant(),'\s*\[[^\]]*\]\s*$','');$t=[regex]::Replace($t,'[.''`]','');$t=[regex]::Replace($t,'[^A-Z0-9, \-]',' ')
  $t=[regex]::Replace($t,'(^|[ ,])(FNP-BC|FNP-C|PA-C|MD|DO|NP|PA|FNP|APRN|ANP|AGNP|AGPCNP|ACNP|CNM|RN|DNP|PHD|MBBS|DPM|JR|SR|II|III|IV)(?=$|[ ,])','$1')
  $t=[regex]::Replace($t,'\s+',' ').Trim();$t=$t.Trim(',',' ');if($t.Length -eq 0){return ''}
  $c=$t.IndexOf(',')
  if($c -ge 0){$last=$t.Substring(0,$c).Trim();$rest=$t.Substring($c+1).Trim(',',' ');$first=$rest;$sp=$rest.IndexOf(' ');if($sp -gt 0){$first=$rest.Substring(0,$sp)};if($first.Length -eq 0){return $last};return $last+', '+$first}
  $w=$t.Split(' ');if($w.Length -eq 1){return $t};return $w[$w.Length-1]+', '+$w[0]
 }
 static [object] Convert([object]$v,[string]$type){
  # Converts to a field type; '' (text) keeps strings, turning empty text into blank.
  if($type -eq 'number'){return [CmbRt]::Num($v)};if($type -eq 'date'){return [CmbRt]::Date($v)};if($type -eq 'bool'){return [CmbRt]::Bool($v)}
  if($type -eq 'text'){if([CmbRt]::IsBlank($v)){return $null};return [CmbRt]::Text($v).Trim()}
  if($v -is [string] -and [string]::IsNullOrWhiteSpace([string]$v)){return $null};return $v
 }
}

# Strict JSON reader (also accepts trailing commas) giving ordered, case-insensitive dictionaries and object[] arrays. Text stays text:
# unlike ConvertFrom-Json in PowerShell 7 it never turns date-like strings into dates, and its errors name the line and column.
class CmbJsonReader{
 [string]$S;[int]$I
 CmbJsonReader([string]$s){$this.S=$s;$this.I=0}
 [void] Fail([string]$m){
  $line=1;$col=1;$end=[Math]::Min($this.I,$this.S.Length);for($k=0;$k -lt $end;$k++){if($this.S[$k] -eq [char]10){$line++;$col=1}else{$col++}}
  throw ('line '+$line+', column '+$col+': '+$m)
 }
 [void] Ws(){while($this.I -lt $this.S.Length -and [char]::IsWhiteSpace($this.S[$this.I])){$this.I++}}
 [object] Document(){
  if($this.S.Length -gt 0 -and $this.S[0] -eq [char]0xFEFF){$this.I=1}
  $this.Ws();if($this.I -ge $this.S.Length){$this.Fail('the text is empty')};$v=$this.Value();$this.Ws()
  if($this.I -lt $this.S.Length){$this.Fail('unexpected text after the end: '+$this.Near())};return $v
 }
 [string] Near(){$n=[Math]::Min(25,$this.S.Length-$this.I);if($n -le 0){return 'end of text'};return '"'+$this.S.Substring($this.I,$n)+'"'}
 [object] Value(){
  $this.Ws();if($this.I -ge $this.S.Length){$this.Fail('unexpected end of text');return $null}
  $c=$this.S[$this.I]
  if($c -eq [char]'{'){return $this.Obj()};if($c -eq [char]'['){return $this.Arr()};if($c -eq [char]'"'){return $this.Str()}
  if($c -eq [char]'-' -or [char]::IsDigit($c)){return $this.Number()}
  if([string]::CompareOrdinal($this.S,$this.I,'true',0,4) -eq 0){$this.I+=4;return $true}
  if([string]::CompareOrdinal($this.S,$this.I,'false',0,5) -eq 0){$this.I+=5;return $false}
  if([string]::CompareOrdinal($this.S,$this.I,'null',0,4) -eq 0){$this.I+=4;return $null}
  if($c -eq [char]0x201C -or $c -eq [char]0x201D){$this.Fail('curly quotes are not JSON; use straight double quotes');return $null}
  $this.Fail('expected a value but found '+$this.Near());return $null
 }
 [object] Obj(){
  $o=[Collections.Specialized.OrderedDictionary]::new([StringComparer]::OrdinalIgnoreCase);$this.I++
  while($true){
   $this.Ws();if($this.I -ge $this.S.Length){$this.Fail('missing } at the end of an object');return $null}
   if($this.S[$this.I] -eq [char]'}'){$this.I++;return $o}
   if($this.S[$this.I] -ne [char]'"'){$this.Fail('expected a "name" in double quotes but found '+$this.Near());return $null}
   $k=$this.Str();$this.Ws();if($this.I -ge $this.S.Length -or $this.S[$this.I] -ne [char]':'){$this.Fail('expected : after "'+$k+'"');return $null}
   $this.I++;$o[$k]=$this.Value();$this.Ws()
   if($this.I -ge $this.S.Length){$this.Fail('missing } at the end of an object');return $null}
   $c=$this.S[$this.I];if($c -eq [char]','){$this.I++;continue};if($c -eq [char]'}'){$this.I++;return $o}
   $this.Fail('expected , or } after the value of "'+$k+'" but found '+$this.Near());return $null
  }
  return $o
 }
 [object] Arr(){
  $l=[Collections.Generic.List[object]]::new();$this.I++
  while($true){
   $this.Ws();if($this.I -ge $this.S.Length){$this.Fail('missing ] at the end of a list');return $null}
   if($this.S[$this.I] -eq [char]']'){$this.I++;return $l.ToArray()}
   $l.Add($this.Value());$this.Ws()
   if($this.I -ge $this.S.Length){$this.Fail('missing ] at the end of a list');return $null}
   $c=$this.S[$this.I];if($c -eq [char]','){$this.I++;continue};if($c -eq [char]']'){$this.I++;return $l.ToArray()}
   $this.Fail('expected , or ] in a list but found '+$this.Near());return $null
  }
  return $l.ToArray()
 }
 [string] Str(){
  $sb=[Text.StringBuilder]::new();$this.I++;$stops=[char[]]@([char]'"',[char]'\',[char]10,[char]13)
  while($true){
   $j=$this.S.IndexOfAny($stops,$this.I);if($j -lt 0){$this.I=$this.S.Length;$this.Fail('text is missing its closing "');return ''}
   [void]$sb.Append($this.S,$this.I,$j-$this.I);$this.I=$j;$c=$this.S[$j]
   if($c -eq [char]'"'){$this.I++;return $sb.ToString()}
   if($c -ne [char]'\'){$this.Fail('text is missing its closing " before the end of the line');return ''}
   if($j+1 -ge $this.S.Length){$this.Fail('text ends after \');return ''}
   $e=$this.S[$j+1];$this.I=$j+2
   switch([string]$e){'"'{[void]$sb.Append('"')}'\'{[void]$sb.Append('\')}'/'{[void]$sb.Append('/')}'b'{[void]$sb.Append([char]8)}'f'{[void]$sb.Append([char]12)}'n'{[void]$sb.Append([char]10)}'r'{[void]$sb.Append([char]13)}'t'{[void]$sb.Append([char]9)}
    'u'{if($this.I+4 -gt $this.S.Length -or $this.S.Substring($this.I,4) -notmatch '^[0-9a-fA-F]{4}$'){$this.Fail('\u needs 4 hex digits');return ''};[void]$sb.Append([char][Convert]::ToInt32($this.S.Substring($this.I,4),16));$this.I+=4}
    default{$this.I=$j;$this.Fail('unknown escape \'+$e+' (write \\ for a backslash)');return ''}}
  }
  return $sb.ToString()
 }
 [object] Number(){
  $m=[regex]::Match($this.S.Substring($this.I,[Math]::Min(64,$this.S.Length-$this.I)),'^-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?')
  if(-not $m.Success){$this.Fail('bad number '+$this.Near());return $null}
  $this.I+=$m.Length;return [double]::Parse($m.Value,[Globalization.CultureInfo]::InvariantCulture)
 }
}
function ConvertFrom-CmbJson([string]$Text){$v=([CmbJsonReader]::new($Text)).Document();if($v -is [Array]){return ,$v};return $v}
function ConvertTo-CmbJsonString([string]$s){
 if($s -notmatch '["\\\x00-\x1f]'){return '"'+$s+'"'}
 $sb=New-Object Text.StringBuilder ($s.Length+16);[void]$sb.Append('"')
 foreach($ch in $s.ToCharArray()){$n=[int]$ch;if($ch -eq '"'){[void]$sb.Append('\"')}elseif($ch -eq '\'){[void]$sb.Append('\\')}elseif($n -eq 10){[void]$sb.Append('\n')}elseif($n -eq 13){[void]$sb.Append('\r')}elseif($n -eq 9){[void]$sb.Append('\t')}elseif($n -lt 32){[void]$sb.Append('\u'+$n.ToString('x4'))}else{[void]$sb.Append($ch)}}
 [void]$sb.Append('"');return $sb.ToString()
}
function ConvertTo-CmbJson($Value,[switch]$SortKeys){
 # Compact JSON that leaves < > & ' unescaped (ConvertTo-Json in Windows PowerShell writes < etc.), keeping key order unless SortKeys.
 if($null -eq $Value){return 'null'}
 if($Value -is [string]){return (ConvertTo-CmbJsonString $Value)}
 if($Value -is [bool]){if($Value){return 'true'};return 'false'}
 if($Value -is [double] -or $Value -is [single]){$d=[double]$Value;if([double]::IsNaN($d) -or [double]::IsInfinity($d)){return 'null'};return $d.ToString('R',[Globalization.CultureInfo]::InvariantCulture)}
 if($Value -is [int] -or $Value -is [long] -or $Value -is [decimal]){return ([string]::Format([Globalization.CultureInfo]::InvariantCulture,'{0}',$Value))}
 if($Value -is [DateTime]){$t=[DateTime]$Value;if($t.TimeOfDay.Ticks -eq 0){return ('"'+$t.ToString('yyyy-MM-dd')+'"')};return ('"'+$t.ToString('yyyy-MM-ddTHH:mm:ss')+'"')}
 if($Value -is [Collections.IDictionary]){
  $keys=@($Value.Keys);if($SortKeys){$keys=@($keys|Sort-Object)}
  $parts=foreach($k in $keys){(ConvertTo-CmbJsonString ([string]$k))+':'+(ConvertTo-CmbJson $Value[$k] -SortKeys:$SortKeys)}
  return '{'+(@($parts) -join ',')+'}'
 }
 if($Value -is [Management.Automation.PSCustomObject]){
  $keys=@($Value.PSObject.Properties|ForEach-Object{$_.Name});if($SortKeys){$keys=@($keys|Sort-Object)}
  $parts=foreach($k in $keys){(ConvertTo-CmbJsonString $k)+':'+(ConvertTo-CmbJson $Value.$k -SortKeys:$SortKeys)}
  return '{'+(@($parts) -join ',')+'}'
 }
 if($Value -is [Collections.IEnumerable]){$parts=foreach($x in $Value){ConvertTo-CmbJson $x -SortKeys:$SortKeys};return '['+(@($parts) -join ',')+']'}
 return (ConvertTo-CmbJsonString ([string]$Value))
}
function Get-CmbValue($Rule,[string]$Key,$Default=$null){if($null -ne $Rule -and $Rule -is [Collections.IDictionary] -and $Rule.Contains($Key) -and $null -ne $Rule[$Key]){return $Rule[$Key]};return $Default}
function Get-CmbText($Rule,[string]$Key){$v=Get-CmbValue $Rule $Key '';if($v -is [string]){return $v.Trim()};if($v -is [double] -or $v -is [bool]){return (ConvertTo-CmbJson $v)};return ''}
function Get-CmbList($Rule,[string]$Key){
 # A property that should be an array of text; a single text value is accepted as a one-item list.
 $v=Get-CmbValue $Rule $Key $null;if($null -eq $v){return @()}
 if($v -is [string]){if($v.Trim()){return @($v.Trim())};return @()}
 if($v -is [Array]){return @(foreach($x in $v){if($null -ne $x -and ([string]$x).Trim()){([string]$x).Trim()}})}
 return @([string]$v)
}
function ConvertTo-CmbRuleText($Rule){
 # One rule as one line: "rule" first and "note" last, the rest in its own order.
 $o=[ordered]@{};if($Rule.Contains('rule')){$o['rule']=$Rule['rule']}
 foreach($k in @($Rule.Keys)){if($k -ne 'rule' -and $k -ne 'note'){$o[[string]$k]=$Rule[$k]}}
 if($Rule.Contains('note')){$o['note']=$Rule['note']}
 return (ConvertTo-CmbJson $o)
}
function Format-CmbConfigText([object[]]$Rules){
 $lines=@(foreach($r in $Rules){ConvertTo-CmbRuleText $r})
 if($lines.Count -eq 0){return "[`r`n]`r`n"}
 return ("[`r`n"+($lines -join (",`r`n"))+"`r`n]`r`n")
}
function Read-CmbConfigText([string]$Text){
 # Accepts a config as saved (a JSON array, one rule per line), pretty-printed JSON, an object holding "rules" or "currentConfig", or rows
 # copied from Excel (one rule per row, with or without the [ ] rows and trailing commas; cells Excel wrapped in quotes are unwrapped).
 $errors=[Collections.Generic.List[object]]::new();$rules=[Collections.Generic.List[object]]::new()
 $t=[string]$Text;if($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF){$t=$t.Substring(1)}
 $lines=@(($t -replace "`r`n","`n" -replace "`r","`n") -split "`n")
 $clean=@(foreach($l in $lines){
  $x=$l.Trim();if($x -match '^```'){'';continue}
  if($x.Length -ge 2 -and $x.StartsWith('"') -and ($x.EndsWith('"') -or $x.EndsWith('",')) -and $x.Contains('""') -and $x.Substring(1).TrimStart().StartsWith('{')){$x=$x.Substring(1);if($x.EndsWith('",')){$x=$x.Substring(0,$x.Length-2)+','}else{$x=$x.Substring(0,$x.Length-1)};$x=$x.Replace('""','"')}
  $x
 })
 $joined=($clean -join "`n").Trim()
 if(!$joined){$errors.Add([ordered]@{line=0;message='Nothing to read: the text is empty.'});return [ordered]@{rules=@();errors=@($errors);lines=@()}}
 $doc=$null;$wholeError=''
 try{$doc=ConvertFrom-CmbJson $joined}catch{$wholeError=$_.Exception.Message}
 if(!$wholeError -and $joined.Contains([string][char]0x201C)){$wholeError=''}
 if(!$wholeError){
  $list=$null
  if($doc -is [Array]){$list=$doc}
  elseif($doc -is [Collections.IDictionary]){foreach($k in @('rules','currentConfig','config')){if($doc.Contains($k) -and $doc[$k] -is [Array]){$list=$doc[$k];break}};if($null -eq $list -and $doc.Contains('rule')){$list=@($doc)}}
  if($null -eq $list){$errors.Add([ordered]@{line=0;message='Expected a list of rules: [ {"rule":...}, ... ].'});return [ordered]@{rules=@();errors=@($errors);lines=@()}}
  $n=0;foreach($r in $list){$n++;if($r -is [Collections.IDictionary]){$rules.Add($r)}else{$errors.Add([ordered]@{line=0;message=('Item '+$n+' of the list is not a rule object.')})}}
  return [ordered]@{rules=@($rules);errors=@($errors);lines=@()}
 }
 # Row by row: each non-empty row holds one rule object.
 $lineNos=[Collections.Generic.List[int]]::new();$rowErrors=0;$rowsTried=0
 for($i=0;$i -lt $clean.Count;$i++){
  $x=[string]$clean[$i];if(!$x -or $x -eq '[' -or $x -eq ']' -or $x -eq '],' -or $x -eq '[,'){continue}
  if($x.StartsWith('[{')){$x=$x.Substring(1)};if($x.EndsWith('}]')){$x=$x.Substring(0,$x.Length-1)};if($x.EndsWith('},')){$x=$x.Substring(0,$x.Length-1)}
  $rowsTried++
  try{$r=ConvertFrom-CmbJson $x;if($r -is [Collections.IDictionary]){$rules.Add($r);$lineNos.Add($i+1)}else{$rowErrors++;$errors.Add([ordered]@{line=($i+1);message='This row is not a rule object.'})}}
  catch{$rowErrors++;$m=$_.Exception.Message -replace '^line 1, ','';$errors.Add([ordered]@{line=($i+1);message=$m})}
 }
 if($rowsTried -eq 0 -or ($rowErrors -gt 0 -and $rules.Count -eq 0) -or $rowErrors -gt [Math]::Max(3,$rowsTried/2)){
  return [ordered]@{rules=@();errors=@([ordered]@{line=0;message=('This is not valid JSON: '+$wholeError)});lines=@()}
 }
 return [ordered]@{rules=@($rules);errors=@($errors);lines=@($lineNos)}
}
function Get-CmbRuleIdentity($Rule,[hashtable]$Seen){
 # A stable name for a rule (used to compare two versions of a config): its "id" if given, otherwise what it applies to.
 $type=([string](Get-CmbValue $Rule 'rule' '')).Trim().ToLowerInvariant();$id=Get-CmbText $Rule 'id'
 if($id){$k='id:'+$id}
 else{
  switch($type){
   'config'{$k='config'}
   'source'{$k='source:'+(Get-CmbText $Rule 'source')}
   'map'{$k='map:'+(Get-CmbText $Rule 'source')+'.'+(Get-CmbText $Rule 'field')}
   'field'{$k='field:'+(Get-CmbText $Rule 'field')}
   'derive'{$k='derive:'+(Get-CmbText $Rule 'field')}
   'lookup'{$k='lookup:'+(Get-CmbText $Rule 'table')}
   'sort'{$k='sort:'+(Get-CmbText $Rule 'field')}
   'output'{$k='output:'+(Get-CmbText $Rule 'field')}
   'view'{$k='view:'+(Get-CmbText $Rule 'view')}
   'measure'{$k='measure:'+(Get-CmbText $Rule 'view')+'.'+(Get-CmbText $Rule 'name')}
   'filter'{$s=Get-CmbText $Rule 'source';if(!$s){$s='patients'};$k='filter:'+$s}
   'key'{$k='key'}
   default{$k='rule:'+$type}
  }
 }
 $k=$k.ToLowerInvariant();if($Seen.ContainsKey($k)){$Seen[$k]++;return ($k+'#'+$Seen[$k])};$Seen[$k]=1;return $k
}
function Compare-CmbRules([object[]]$Old,[object[]]$New){
 # What changed between two versions: added, removed and changed rules (ignoring key order inside a rule), and order changes where order matters.
 $so=@{};$sn=@{};$a=[ordered]@{};$b=[ordered]@{}
 foreach($r in @($Old)){$a[(Get-CmbRuleIdentity $r $so)]=$r};foreach($r in @($New)){$b[(Get-CmbRuleIdentity $r $sn)]=$r}
 $out=[Collections.Generic.List[object]]::new()
 foreach($k in @($b.Keys)){
  $after=ConvertTo-CmbRuleText $b[$k]
  if(!$a.Contains($k)){$out.Add([ordered]@{change='added';identity=$k;before='';after=$after});continue}
  if((ConvertTo-CmbJson $a[$k] -SortKeys) -ne (ConvertTo-CmbJson $b[$k] -SortKeys)){$out.Add([ordered]@{change='changed';identity=$k;before=(ConvertTo-CmbRuleText $a[$k]);after=$after})}
 }
 foreach($k in @($a.Keys)){if(!$b.Contains($k)){$out.Add([ordered]@{change='removed';identity=$k;before=(ConvertTo-CmbRuleText $a[$k]);after=''})}}
 foreach($type in @('source','key','derive','sort','output','measure')){
  $oa=@($a.Keys|Where-Object{$_ -like ($type+'*') -and $b.Contains($_)});$ob=@($b.Keys|Where-Object{$_ -like ($type+'*') -and $a.Contains($_)})
  if(($oa -join '|') -ne ($ob -join '|')){$out.Add([ordered]@{change='reordered';identity=$type;before=($oa -join ', ');after=($ob -join ', ')})}
 }
 return ,@($out)
}

# --- Combine expressions: tokenized, parsed and compiled into one PowerShell loop per expression. User text never becomes code: names
# become column variables, text and numbers become constants, and functions and operators come from a fixed list.
$script:CmbFunctions=@{
 'if'=@(3,3);'coalesce'=@(1,20);'isblank'=@(1,1);'upper'=@(1,1);'lower'=@(1,1);'trim'=@(1,1);'len'=@(1,1);'left'=@(2,2);'right'=@(2,2);'mid'=@(3,3)
 'contains'=@(2,2);'startswith'=@(2,2);'endswith'=@(2,2);'matches'=@(2,2);'replace'=@(3,3);'extract'=@(2,2);'number'=@(1,1);'date'=@(1,1);'text'=@(1,1)
 'round'=@(1,2);'abs'=@(1,1);'min'=@(1,20);'max'=@(1,20);'today'=@(0,0);'year'=@(1,1);'month'=@(1,1);'day'=@(1,1);'adddays'=@(2,2);'addmonths'=@(2,2);'addyears'=@(2,2)
 'daysbetween'=@(2,2);'monthsbetween'=@(2,2);'yearsbetween'=@(2,2);'format'=@(2,2);'lookup'=@(2,3);'join'=@(2,60);'latestdate'=@(1,2);'countitems'=@(1,3)
}
$script:CmbBoolFunctions=@('isblank','contains','startswith','endswith','matches')
function Get-CmbTokens([string]$Text){
 $t=[Collections.Generic.List[object]]::new();$i=0;$n=$Text.Length
 while($i -lt $n){
  $c=$Text[$i]
  if([char]::IsWhiteSpace($c)){$i++;continue}
  if($c -eq "'" -or $c -eq '"'){
   $q=$c;$sb=New-Object Text.StringBuilder;$j=$i+1;$closed=$false
   while($j -lt $n){if($Text[$j] -eq $q){if($j+1 -lt $n -and $Text[$j+1] -eq $q){[void]$sb.Append($q);$j+=2;continue};$closed=$true;break};[void]$sb.Append($Text[$j]);$j++}
   if(!$closed){throw ('text starting at position '+($i+1)+' has no closing quote')}
   $t.Add(@{t='str';v=$sb.ToString();p=$i});$i=$j+1;continue
  }
  if($c -eq '['){$j=$Text.IndexOf(']',$i+1);if($j -lt 0){throw ('[ at position '+($i+1)+' has no closing ]')};$name=$Text.Substring($i+1,$j-$i-1).Trim();if(!$name){throw ('empty [] at position '+($i+1))};$t.Add(@{t='id';v=$name;p=$i;br=$true});$i=$j+1;continue}
  if([char]::IsDigit($c) -or ($c -eq '.' -and $i+1 -lt $n -and [char]::IsDigit($Text[$i+1]))){$m=[regex]::Match($Text.Substring($i),'^\d*\.?\d+([eE][+-]?\d+)?');$t.Add(@{t='num';v=[double]::Parse($m.Value,[Globalization.CultureInfo]::InvariantCulture);p=$i});$i+=$m.Length;continue}
  if([char]::IsLetter($c) -or $c -eq '_'){
   $m=[regex]::Match($Text.Substring($i),'^[A-Za-z_][A-Za-z0-9_]*');$w=$m.Value;$lw=$w.ToLowerInvariant()
   if($lw -in @('and','or','not','in','is','true','false','blank','null')){$t.Add(@{t='kw';v=$lw;p=$i})}else{$t.Add(@{t='id';v=$w;p=$i;br=$false})}
   $i+=$m.Length;continue
  }
  $two='';if($i+1 -lt $n){$two=$Text.Substring($i,2)}
  if($two -in @('<=','>=','<>','!=','==')){$t.Add(@{t='op';v=$two;p=$i});$i+=2;continue}
  if('=<>&+-*/'.IndexOf($c) -ge 0){$t.Add(@{t='op';v=[string]$c;p=$i});$i++;continue}
  if($c -eq '('){$t.Add(@{t='lp';v='(';p=$i});$i++;continue};if($c -eq ')'){$t.Add(@{t='rp';v=')';p=$i});$i++;continue};if($c -eq ','){$t.Add(@{t='comma';v=',';p=$i});$i++;continue}
  if($c -eq [char]0x2018 -or $c -eq [char]0x2019 -or $c -eq [char]0x201C -or $c -eq [char]0x201D){throw ('curly quote at position '+($i+1)+'; use a straight single quote')}
  throw ('unexpected character "'+$c+'" at position '+($i+1))
 }
 $t.Add(@{t='end';v='';p=$n});return ,$t
}
class CmbExprParser{
 [object]$T;[int]$K
 CmbExprParser([object]$tokens){$this.T=$tokens;$this.K=0}
 [hashtable] Peek(){return [hashtable]$this.T[$this.K]}
 [hashtable] Next(){$x=[hashtable]$this.T[$this.K];if($this.K -lt $this.T.Count-1){$this.K++};return $x}
 [bool] IsKw([string]$w){$x=$this.Peek();return ($x.t -eq 'kw' -and $x.v -eq $w)}
 [bool] IsOp([string]$o){$x=$this.Peek();return ($x.t -eq 'op' -and $x.v -eq $o)}
 [void] Fail([string]$m){$x=$this.Peek();$where='at the end';if($x.t -ne 'end'){$where='at position '+($x.p+1)};throw ($m+' '+$where)}
 [hashtable] Parse(){$e=$this.OrX();if($this.Peek().t -ne 'end'){$this.Fail('unexpected "'+$this.Peek().v+'"')};return $e}
 [hashtable] OrX(){$l=$this.AndX();while($this.IsKw('or')){$null=$this.Next();$l=@{t='or';l=$l;r=$this.AndX()}};return $l}
 [hashtable] AndX(){$l=$this.NotX();while($this.IsKw('and')){$null=$this.Next();$l=@{t='and';l=$l;r=$this.NotX()}};return $l}
 [hashtable] NotX(){if($this.IsKw('not')){$null=$this.Next();return @{t='not';e=$this.NotX()}};return $this.CmpX()}
 [hashtable] CmpX(){
  $l=$this.CatX();$x=$this.Peek()
  if($x.t -eq 'op' -and $x.v -in @('=','==','<>','!=','<','<=','>','>=')){$null=$this.Next();$op=$x.v;if($op -eq '=='){$op='='};if($op -eq '!='){$op='<>'};return @{t='cmp';op=$op;l=$l;r=$this.CatX()}}
  $neg=$false
  if($this.IsKw('not')){$save=$this.K;$null=$this.Next();if($this.IsKw('in')){$neg=$true}else{$this.K=$save;return $l}}
  if($this.IsKw('in')){
   $null=$this.Next();if($this.Peek().t -ne 'lp'){$this.Fail('expected ( after in')};$null=$this.Next()
   $items=[Collections.Generic.List[object]]::new();if($this.Peek().t -ne 'rp'){$items.Add($this.CatX());while($this.Peek().t -eq 'comma'){$null=$this.Next();$items.Add($this.CatX())}}
   if($this.Peek().t -ne 'rp'){$this.Fail('expected ) to close the in list')};$null=$this.Next();return @{t='in';e=$l;items=$items.ToArray();neg=$neg}
  }
  if($this.IsKw('is')){$null=$this.Next();$n=$false;if($this.IsKw('not')){$null=$this.Next();$n=$true};if(-not ($this.IsKw('blank') -or $this.IsKw('null'))){$this.Fail('expected blank after is')};$null=$this.Next();return @{t='isblank';e=$l;neg=$n}}
  return $l
 }
 [hashtable] CatX(){$l=$this.AddX();while($this.IsOp('&')){$null=$this.Next();$l=@{t='bin';op='&';l=$l;r=$this.AddX()}};return $l}
 [hashtable] AddX(){$l=$this.MulX();while($this.IsOp('+') -or $this.IsOp('-')){$op=$this.Next().v;$l=@{t='bin';op=$op;l=$l;r=$this.MulX()}};return $l}
 [hashtable] MulX(){$l=$this.UnX();while($this.IsOp('*') -or $this.IsOp('/')){$op=$this.Next().v;$l=@{t='bin';op=$op;l=$l;r=$this.UnX()}};return $l}
 [hashtable] UnX(){if($this.IsOp('-')){$null=$this.Next();return @{t='neg';e=$this.UnX()}};return $this.Prim()}
 [hashtable] Prim(){
  $x=$this.Next()
  if($x.t -eq 'num'){return @{t='const';v=$x.v}};if($x.t -eq 'str'){return @{t='const';v=$x.v}}
  if($x.t -eq 'kw'){if($x.v -eq 'true'){return @{t='const';v=$true}};if($x.v -eq 'false'){return @{t='const';v=$false}};if($x.v -eq 'blank' -or $x.v -eq 'null'){return @{t='null'}};$this.K--;$this.Fail('unexpected "'+$x.v+'"')}
  if($x.t -eq 'lp'){$e=$this.OrX();if($this.Peek().t -ne 'rp'){$this.Fail('expected )')};$null=$this.Next();return $e}
  if($x.t -eq 'id'){
   if(-not $x.br -and $this.Peek().t -eq 'lp'){
    $null=$this.Next();$args1=[Collections.Generic.List[object]]::new()
    if($this.Peek().t -ne 'rp'){$args1.Add($this.OrX());while($this.Peek().t -eq 'comma'){$null=$this.Next();$args1.Add($this.OrX())}}
    if($this.Peek().t -ne 'rp'){$this.Fail('expected ) to close '+$x.v+'(')};$null=$this.Next();return @{t='call';name=([string]$x.v).ToLowerInvariant();shown=$x.v;args=$args1.ToArray();p=$x.p}
   }
   return @{t='id';name=[string]$x.v;p=$x.p}
  }
  if($x.t -eq 'end'){$this.Fail('the expression is incomplete')}
  $this.K--;$this.Fail('unexpected "'+$x.v+'"');return $null
 }
}
function ConvertTo-CmbExprTree([string]$Text){
 if([string]::IsNullOrWhiteSpace($Text)){throw 'the expression is empty'}
 return ([CmbExprParser]::new((Get-CmbTokens $Text))).Parse()
}
function New-CmbExprContext([hashtable]$Names,[hashtable]$Tables,$Today,[string]$Unknown='',[switch]$AcceptAny){
 # Names maps a name (ConvertTo-CmbHeaderKey form) to the index of a column array; AcceptAny lets unknown names through (headers not known
 # yet). Tables maps a lookup table name (lower case) to @{index;fields}.
 return @{Names=$Names;AcceptAny=[bool]$AcceptAny;Unknown=$Unknown;Tables=$Tables;Today=$Today;Consts=[Collections.Generic.List[object]]::new();Cols=[Collections.Generic.List[int]]::new();ColVar=@{};Refs=[Collections.Generic.List[string]]::new();TableRefs=[Collections.Generic.List[int]]::new()}
}
function Add-CmbConst($Ctx,$Value){$Ctx.Consts.Add($Value);return ('$k'+($Ctx.Consts.Count-1))}
function ConvertTo-CmbCode($Node,$Ctx){
 # Returns @{code;bool}: PowerShell for one row (row index $i), and whether it already gives [bool].
 switch($Node.t){
  'const'{return @{code=(Add-CmbConst $Ctx $Node.v);bool=($Node.v -is [bool])}}
  'null'{return @{code='$null';bool=$false}}
  'id'{
   $nk=ConvertTo-CmbHeaderKey ([string]$Node.name)
   if($Ctx.Names.ContainsKey($nk)){$col=[int]$Ctx.Names[$nk]}elseif($Ctx.AcceptAny){$col=0}else{throw ('unknown name "'+$Node.name+'"'+$(if($Ctx.Unknown){' ('+$Ctx.Unknown+')'}else{''})+' at position '+($Node.p+1))}
   $Ctx.Refs.Add([string]$Node.name)
   if(!$Ctx.ColVar.ContainsKey($col)){$Ctx.Cols.Add($col);$Ctx.ColVar[$col]='$c'+($Ctx.Cols.Count-1)}
   return @{code=($Ctx.ColVar[$col]+'[$i]');bool=$false}
  }
  'not'{$e=ConvertTo-CmbCode $Node.e $Ctx;return @{code=('(-not '+(Get-CmbBoolCode $e)+')');bool=$true}}
  'and'{$l=ConvertTo-CmbCode $Node.l $Ctx;$r=ConvertTo-CmbCode $Node.r $Ctx;return @{code=('('+(Get-CmbBoolCode $l)+' -and '+(Get-CmbBoolCode $r)+')');bool=$true}}
  'or'{$l=ConvertTo-CmbCode $Node.l $Ctx;$r=ConvertTo-CmbCode $Node.r $Ctx;return @{code=('('+(Get-CmbBoolCode $l)+' -or '+(Get-CmbBoolCode $r)+')');bool=$true}}
  'cmp'{
   $l=(ConvertTo-CmbCode $Node.l $Ctx).code;$r=(ConvertTo-CmbCode $Node.r $Ctx).code
   $m=@{'='='Eq';'<>'='Eq';'<'='Lt';'<='='Le';'>'='Gt';'>='='Ge'}[$Node.op];$c='[CmbRt]::'+$m+'(('+$l+'),('+$r+'))'
   if($Node.op -eq '<>'){$c='(-not '+$c+')'};return @{code=$c;bool=$true}
  }
  'in'{
   $e=(ConvertTo-CmbCode $Node.e $Ctx).code;$allConst=$true;foreach($x in $Node.items){if($x.t -ne 'const'){$allConst=$false}}
   if($allConst){$list=Add-CmbConst $Ctx ([object[]]@(foreach($x in $Node.items){$x.v}))}else{$list='[object[]]@('+(@(foreach($x in $Node.items){'('+(ConvertTo-CmbCode $x $Ctx).code+')'}) -join ',')+')'}
   $c='[CmbRt]::In(('+$e+'),'+$list+')';if($Node.neg){$c='(-not '+$c+')'};return @{code=$c;bool=$true}
  }
  'isblank'{$e=(ConvertTo-CmbCode $Node.e $Ctx).code;$c='[CmbRt]::IsBlank(('+$e+'))';if($Node.neg){$c='(-not '+$c+')'};return @{code=$c;bool=$true}}
  'neg'{$e=(ConvertTo-CmbCode $Node.e $Ctx).code;return @{code=('[CmbRt]::Neg(('+$e+'))');bool=$false}}
  'bin'{
   $l=(ConvertTo-CmbCode $Node.l $Ctx).code;$r=(ConvertTo-CmbCode $Node.r $Ctx).code
   $m=@{'&'='Cat';'+'='Add';'-'='Sub';'*'='Mul';'/'='Div'}[$Node.op];return @{code=('[CmbRt]::'+$m+'(('+$l+'),('+$r+'))');bool=$false}
  }
  'call'{return (ConvertTo-CmbCallCode $Node $Ctx)}
 }
 throw ('unknown expression part '+$Node.t)
}
function Get-CmbBoolCode($Part){if($Part.bool){return $Part.code};return ('[CmbRt]::Truthy(('+$Part.code+'))')}
function ConvertTo-CmbCallCode($Node,$Ctx){
 $name=[string]$Node.name;$a=@($Node.args)
 if(!$script:CmbFunctions.ContainsKey($name)){throw ('unknown function '+$Node.shown+'() at position '+($Node.p+1))}
 $lim=$script:CmbFunctions[$name];if($a.Count -lt $lim[0] -or $a.Count -gt $lim[1]){$want=$(if($lim[0] -eq $lim[1]){[string]$lim[0]}else{[string]$lim[0]+' to '+$lim[1]});throw ($Node.shown+'() takes '+$want+' value'+$(if($lim[1] -eq 1){''}else{'s'})+' but was given '+$a.Count)}
 if($name -eq 'today'){return @{code=(Add-CmbConst $Ctx $Ctx.Today);bool=$false}}
 if($name -eq 'lookup'){
  if($a[0].t -ne 'const' -or !($a[0].v -is [string])){throw 'lookup() needs the table name in quotes as its first value'}
  $tn=([string]$a[0].v).ToLowerInvariant();if(!$Ctx.Tables.ContainsKey($tn)){throw ('lookup(): there is no lookup table or lookup source named "'+$a[0].v+'"')}
  $tb=$Ctx.Tables[$tn];$field=''
  if($a.Count -eq 3){if($a[2].t -ne 'const' -or !($a[2].v -is [string])){throw 'lookup() needs the field name in quotes as its third value'};$field=[string]$a[2].v;$hit=@($tb.fields|Where-Object{$_ -ieq $field});if($tb.fields.Count -gt 0 -and $hit.Count -eq 0){throw ('lookup(): source "'+$a[0].v+'" has no field "'+$field+'"')};if($hit.Count){$field=[string]$hit[0]}}
  elseif($tb.fields.Count -gt 0){throw ('lookup(): "'+$a[0].v+'" is a lookup source; name the field to return, e.g. lookup('''+$a[0].v+''', x, '''+$tb.fields[0]+''')')}
  $Ctx.TableRefs.Add([int]$tb.index)
  return @{code=('[CmbRt]::Lookup($L['+[int]$tb.index+'],('+(ConvertTo-CmbCode $a[1] $Ctx).code+'),'+(Add-CmbConst $Ctx $field)+')');bool=$false}
 }
 if($name -eq 'if'){$c=ConvertTo-CmbCode $a[0] $Ctx;$x=(ConvertTo-CmbCode $a[1] $Ctx).code;$y=(ConvertTo-CmbCode $a[2] $Ctx).code;return @{code=('$(if('+(Get-CmbBoolCode $c)+'){'+$x+'}else{'+$y+'})');bool=$false}}
 $codes=@(foreach($x in $a){'('+(ConvertTo-CmbCode $x $Ctx).code+')'})
 switch($name){
  'coalesce'{return @{code=('[CmbRt]::Coalesce([object[]]@('+($codes -join ',')+'))');bool=$false}}
  'join'{return @{code=('[CmbRt]::JoinList('+$codes[0]+',[object[]]@('+(@($codes|Select-Object -Skip 1) -join ',')+'))');bool=$false}}
  'latestdate'{$u='$null';if($codes.Count -gt 1){$u=$codes[1]};return @{code=('[CmbRt]::LatestDate('+$codes[0]+','+$u+')');bool=$false}}
  'countitems'{$sp=(Add-CmbConst $Ctx ',');$ig='$null';if($codes.Count -gt 1){$sp=$codes[1]};if($codes.Count -gt 2){$ig=$codes[2]};return @{code=('[CmbRt]::CountItems('+$codes[0]+','+$sp+','+$ig+')');bool=$false}}
  'min'{return @{code=('[CmbRt]::Min([object[]]@('+($codes -join ',')+'))');bool=$false}}
  'max'{return @{code=('[CmbRt]::Max([object[]]@('+($codes -join ',')+'))');bool=$false}}
  'isblank'{return @{code=('[CmbRt]::IsBlank('+$codes[0]+')');bool=$true}}
  'number'{return @{code=('[CmbRt]::Num('+$codes[0]+')');bool=$false}}
  'date'{return @{code=('[CmbRt]::Date('+$codes[0]+')');bool=$false}}
  'text'{return @{code=('[CmbRt]::Blank([CmbRt]::Text('+$codes[0]+'))');bool=$false}}
  'round'{$d='0';if($codes.Count -gt 1){$d=$codes[1]};return @{code=('[CmbRt]::Round('+$codes[0]+','+$d+')');bool=$false}}
  'year'{return @{code=('[CmbRt]::Part('+$codes[0]+',''y'')');bool=$false}}
  'month'{return @{code=('[CmbRt]::Part('+$codes[0]+',''m'')');bool=$false}}
  'day'{return @{code=('[CmbRt]::Part('+$codes[0]+',''d'')');bool=$false}}
  'adddays'{return @{code=('[CmbRt]::AddTo('+$codes[0]+','+$codes[1]+',''d'')');bool=$false}}
  'addmonths'{return @{code=('[CmbRt]::AddTo('+$codes[0]+','+$codes[1]+',''m'')');bool=$false}}
  'addyears'{return @{code=('[CmbRt]::AddTo('+$codes[0]+','+$codes[1]+',''y'')');bool=$false}}
  'daysbetween'{return @{code=('[CmbRt]::Between('+$codes[0]+','+$codes[1]+',''d'')');bool=$false}}
  'monthsbetween'{return @{code=('[CmbRt]::Between('+$codes[0]+','+$codes[1]+',''m'')');bool=$false}}
  'yearsbetween'{return @{code=('[CmbRt]::Between('+$codes[0]+','+$codes[1]+',''y'')');bool=$false}}
 }
 $method=@{'upper'='Upper';'lower'='Lower';'trim'='Trim';'len'='Len';'left'='Left';'right'='Right';'mid'='Mid';'contains'='Contains';'startswith'='StartsWith';'endswith'='EndsWith';'matches'='Matches';'replace'='Replace';'extract'='Extract';'abs'='Abs';'format'='Format'}[$name]
 if($name -in @('matches','replace','extract') -and $a[1].t -eq 'const'){try{$null=[regex]::new([string]$a[1].v)}catch{throw ($Node.shown+'(): the pattern '''+$a[1].v+''' is not a valid regular expression')}}
 return @{code=('[CmbRt]::'+$method+'('+($codes -join ',')+')');bool=($script:CmbBoolFunctions -contains $name)}
}
function New-CmbCompiled([string]$Text,$Ctx,[switch]$AsBool){
 # Compiles one expression into a loop over $n rows: & $sb $Cols $Consts $Tables $n $out $errs. Errors name the expression.
 try{$tree=ConvertTo-CmbExprTree $Text;$part=ConvertTo-CmbCode $tree $Ctx}catch{throw ('expression "'+$Text+'": '+$_.Exception.Message)}
 $body=$(if($AsBool){Get-CmbBoolCode $part}else{$part.code})
 $sb=New-Object Text.StringBuilder;[void]$sb.Append('param($C,$K,$L,[int]$n,$out,$errs)')
 for($j=0;$j -lt $Ctx.Cols.Count;$j++){[void]$sb.Append(';$c'+$j+'=$C['+$j+']')}
 for($j=0;$j -lt $Ctx.Consts.Count;$j++){[void]$sb.Append(';$k'+$j+'=$K['+$j+']')}
 [void]$sb.Append(';for($i=0;$i -lt $n;$i++){try{$out[$i]=('+$body+')}catch{$errs.Add($i)}}')
 return @{text=$Text;script=[scriptblock]::Create($sb.ToString());cols=@($Ctx.Cols);consts=$Ctx.Consts.ToArray();refs=@($Ctx.Refs);tables=@($Ctx.TableRefs|Select-Object -Unique);bool=[bool]$AsBool}
}
function Invoke-CmbCompiled($Compiled,[object[]]$Columns,[object[]]$Tables,[int]$Count){
 # Runs a compiled expression over every row; returns @{values;errors(row indexes)}.
 $cols=[object[]]@(foreach($j in $Compiled.cols){,$Columns[$j]})
 $out=New-Object object[] $Count;$errs=[Collections.Generic.List[int]]::new()
 & $Compiled.script $cols $Compiled.consts $Tables $Count $out $errs
 return @{values=$out;errors=$errs}
}
# --- Combine files: the loaded exports (combine\source-files, newest few kept, one record each in state\cmbfile-<token>.json) and their fingerprints.
function Get-CmbFilePath([string]$Token){if($Token -notmatch '^[a-f0-9]{32}$'){throw 'Invalid file token.'};return (Join-Path $script:Paths.State ('cmbfile-'+$Token+'.json'))}
function Get-CmbFiles{return @(Get-ChildItem $script:Paths.State -Filter 'cmbfile-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Where-Object{$null -ne $_}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'loadedUtc' '')}} -Descending)}
function Get-CmbFile([string]$Token){$f=Json (Get-CmbFilePath $Token);if(!$f){throw 'That file is no longer loaded; load it again on the Files page.'};if(!(Test-Path -LiteralPath ([string](Get-P $f 'path' '')))){throw ('"'+[string]$f.originalName+'" was removed from the combine folder; load it again.')};return $f}
function ConvertTo-CmbHeaderKey([string]$Text){return (($Text -replace '\s+',' ').Trim().ToLowerInvariant())}
function Get-CmbShape([string]$Value){
 # The shape of a value for fingerprints: digits become 9, letter runs become A, the rest is kept ("123-45-6789" -> "999-99-9999").
 $s=$Value;if($s.Length -gt 40){$s=$s.Substring(0,40)+'...'};$s=[regex]::Replace($s,'[0-9]','9');$s=[regex]::Replace($s,'[A-Za-z]+','A');return [regex]::Replace($s,'\s+',' ')
}
function Get-CmbValueKind([string]$Value){
 if($Value -match '^-?\d+$'){return 'integer'};if($Value -match '^[-(]?\$?-?[\d,]*\.?\d+%?\)?$'){return 'number'}
 if($Value -match '^(?i:y|n|yes|no|true|false)$'){return 'bool'}
 if($Value -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}' -or $Value -match '^\d{1,4}[-/.]\d{1,2}[-/.]\d{2,4}( \d{1,2}:\d{2}(:\d{2})?( ?[AaPp][Mm])?)?$' -or $Value -match '^\d{1,2}[- ][A-Za-z]{3}[- ]\d{2,4}$' -or $Value -match '^[A-Za-z]{3,9} \d{1,2},? \d{4}$'){if($null -ne [CmbRt]::Date($Value)){return 'date'}}
 return 'text'
}
function Get-CmbTextEncoding([byte[]]$Bytes){
 if($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF){return (New-Object Text.UTF8Encoding($false))}
 if($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE){return [Text.Encoding]::Unicode};if($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF){return [Text.Encoding]::BigEndianUnicode}
 try{$null=(New-Object Text.UTF8Encoding($false,$true)).GetString($Bytes);return (New-Object Text.UTF8Encoding($false))}catch{}
 try{return [Text.Encoding]::GetEncoding(1252)}catch{return [Text.Encoding]::GetEncoding(28591)}
}
function Read-CmbTextRows([string]$Path){
 # A delimited text export as rows of fields. The delimiter is the one of , tab | ; seen most in the first line (tab for .tsv).
 $bytes=[IO.File]::ReadAllBytes($Path);$enc=Get-CmbTextEncoding $bytes;$text=$enc.GetString($bytes);if($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF){$text=$text.Substring(1)}
 $first='';foreach($l in ($text -split "`n",30)){if($l.Trim()){$first=$l;break}}
 $delim=',';$best=-1;foreach($d in @(',',"`t",'|',';')){$n=([regex]::Matches(($first -replace '"[^"]*"',''),[regex]::Escape($d))).Count;if($n -gt $best){$best=$n;$delim=$d}}
 if([IO.Path]::GetExtension($Path) -ieq '.tsv'){$delim="`t"}
 try{Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop}catch{}
 $p=New-Object Microsoft.VisualBasic.FileIO.TextFieldParser (New-Object IO.StringReader $text);$p.SetDelimiters([string[]]@($delim));$p.HasFieldsEnclosedInQuotes=$true;$p.TrimWhiteSpace=$false
 $rows=[Collections.Generic.List[string[]]]::new();$bad=0
 try{while(-not $p.EndOfData){try{$f=$p.ReadFields();if($null -ne $f){$rows.Add($f)}}catch [Microsoft.VisualBasic.FileIO.MalformedLineException]{$bad++;$rows.Add([string[]]@(($p.ErrorLine -split [regex]::Escape($delim))))}}}finally{$p.Close()}
 return @{rows=$rows;delimiter=$delim;encoding=$enc.WebName;malformed=$bad}
}
function Find-CmbHeaderRow([object[]]$Rows){
 # Index (0-based) of the header row among the first rows: the one with the most non-blank, non-numeric cells (ties: the earliest).
 $best=0;$bestN=-1
 for($r=0;$r -lt [Math]::Min(25,$Rows.Count);$r++){$n=0;foreach($v in @($Rows[$r])){$s=([string]$v).Trim();if($s -and $s -notmatch '^-?[\d.,/:\- ]+$'){$n++}};if($n -gt $bestN){$best=$r;$bestN=$n}}
 return $best
}
function Get-CmbUniqueHeaders([string[]]$Headers){
 # Blank headers become "Column C"; repeated headers get " (2)", " (3)" so every column has its own name.
 $seen=@{};$out=New-Object string[] $Headers.Count
 for($c=0;$c -lt $Headers.Count;$c++){
  $h=([string]$Headers[$c] -replace '\s+',' ').Trim();if(!$h){$h='Column '+(Get-ColumnLetter ($c+1))}
  $k=$h.ToLowerInvariant();if($seen.ContainsKey($k)){$seen[$k]++;$h=$h+' ('+$seen[$k]+')'}else{$seen[$k]=1};$out[$c]=$h
 }
 return ,$out
}
function Read-CmbSheets([string]$Path,[string]$OnlySheet='',[int]$HeaderRow=0,[string[]]$Want=$null,[string]$TextName='',[switch]$CellText){
 # Every non-empty table in a file as @{name;headerRow;headers;columns (header -> string[]);rows}. Want limits which columns are read
 # (header keys); HeaderRow (1-based) overrides the detected header row. Text files have one table named after the file.
 $ext=[IO.Path]::GetExtension($Path).ToLowerInvariant();$out=[Collections.Generic.List[object]]::new()
 $wantSet=$null;if($null -ne $Want){$wantSet=@{};foreach($w in $Want){$wantSet[(ConvertTo-CmbHeaderKey $w)]=$true}}
 if($ext -eq '.xlsx'){
  $pkg=$null
  try{
   $pkg=Open-ExcelPackage -Path $Path -ErrorAction Stop
   foreach($ws in $pkg.Workbook.Worksheets){
    if($OnlySheet -and $ws.Name -ne $OnlySheet){continue};if($null -eq $ws.Dimension){continue}
    $maxCol=[Math]::Min(500,$ws.Dimension.End.Column);$last=$ws.Dimension.End.Row;$top=[Math]::Min(25,$last)
    $raw=$ws.Cells[1,1,$top,$maxCol].Value;$grid=@(for($r=0;$r -lt $top;$r++){,([object[]]@(for($c=0;$c -lt $maxCol;$c++){if($raw -is [Array] -and $raw.Rank -eq 2){$raw[$r,$c]}elseif($r -eq 0 -and $c -eq 0){$raw}else{$null}}))})
    $h=$(if($HeaderRow -gt 0){$HeaderRow-1}else{Find-CmbHeaderRow $grid});if($h -ge $grid.Count){continue}
    $headers=Get-CmbUniqueHeaders ([string[]]@(foreach($v in $grid[$h]){if($v -is [DateTime]){([DateTime]$v).ToString('yyyy-MM-dd')}else{[string]$v}}))
    $width=$headers.Count;while($width -gt 0 -and [string]$grid[$h][$width-1] -eq ''){$width--};if($width -eq 0){continue}
    $cols=[ordered]@{};$first=$h+2
    for($c=0;$c -lt $width;$c++){$name=$headers[$c];if($null -ne $wantSet -and !$wantSet.ContainsKey((ConvertTo-CmbHeaderKey $name))){continue};if($last -ge $first){if($CellText){$cols[$name]=Get-CmbCellTextColumn $ws ($c+1) $first $last}else{$cols[$name]=ConvertTo-CmbDateColumn $ws ($c+1) $first (Get-ColumnText $ws ($c+1) $first $last)}}else{$cols[$name]=New-Object string[] 0}}
    $n=[Math]::Max(0,$last-$first+1);$n=Get-CmbTrimmedCount $cols $n
    $out.Add(@{name=$ws.Name;headerRow=($h+1);headers=[string[]]$headers[0..($width-1)];columns=$cols;rows=$n})
   }
  }finally{if($pkg){Close-ExcelPackage $pkg -NoSave}}
  return @($out)
 }
 if($ext -notin @('.csv','.txt','.tsv')){throw ('Unsupported file type '+$ext+'; use .xlsx or .csv (save .xls files as .xlsx first).')}
 $t=Read-CmbTextRows $Path;$rows=$t.rows;if($rows.Count -eq 0){return @()}
 $h=$(if($HeaderRow -gt 0){$HeaderRow-1}else{Find-CmbHeaderRow @($rows.GetRange(0,[Math]::Min(25,$rows.Count)))})
 if($h -ge $rows.Count){return @()}
 $headers=Get-CmbUniqueHeaders ([string[]]$rows[$h]);$width=$headers.Count;$n=$rows.Count-$h-1;$cols=[ordered]@{}
 for($c=0;$c -lt $width;$c++){
  $name=$headers[$c];if($null -ne $wantSet -and !$wantSet.ContainsKey((ConvertTo-CmbHeaderKey $name))){continue}
  $a=New-Object string[] ([Math]::Max(0,$n));for($r=0;$r -lt $n;$r++){$f=$rows[$h+1+$r];if($c -lt $f.Length){$a[$r]=$f[$c].Trim()}else{$a[$r]=''}};$cols[$name]=$a
 }
 $n=Get-CmbTrimmedCount $cols ([Math]::Max(0,$n))
 if(!$TextName){$TextName=[IO.Path]::GetFileNameWithoutExtension($Path)}
 $out.Add(@{name=$TextName;headerRow=($h+1);headers=$headers;columns=$cols;rows=$n;delimiter=$t.delimiter;encoding=$t.encoding;malformed=$t.malformed})
 return @($out)
}
function Get-CmbCellTextColumn($Worksheet,[int]$Column,[int]$FromRow,[int]$ToRow){
 # What Excel shows in each cell, read the way the provider reports read their columns: real dates as ISO text, everything else as displayed.
 $n=$ToRow-$FromRow+1;$out=[string[]]::new([Math]::Max(0,$n))
 for($i=0;$i -lt $n;$i++){$cell=$Worksheet.Cells[($FromRow+$i),$Column];$v=$cell.Value;if($null -eq $v){$out[$i]='';continue};if($v -is [DateTime]){$out[$i]=([DateTime]$v).ToString('o')}else{$out[$i]=([string]$cell.Text).Trim()}}
 return ,$out
}
function ConvertTo-CmbDateColumn($Worksheet,[int]$Column,[int]$FirstRow,[string[]]$Values){
 # Dates saved with a custom number format come back as day numbers; when the column's first value is a number shown as a date, every
 # number in the column becomes the date text Get-ColumnText gives for real dates.
 $k=-1;for($i=0;$i -lt [Math]::Min(200,$Values.Length);$i++){if($Values[$i] -ne ''){$k=$i;break}};if($k -lt 0){return ,$Values}
 if($Values[$k] -match '^\d{4}-\d\d-\d\dT'){
  # Real dates: midnight timestamps are shown as plain dates (2026-03-01), as in the export.
  $rx=[regex]'^(\d{4}-\d\d-\d\d)T00:00:00(\.0+)?$';for($i=0;$i -lt $Values.Length;$i++){$m=$rx.Match($Values[$i]);if($m.Success){$Values[$i]=$m.Groups[1].Value}};return ,$Values
 }
 if($Values[$k] -notmatch '^\d+(\.\d+)?$'){return ,$Values}
 $fmt=[string]$Worksheet.Cells[($FirstRow+$k),$Column].Style.Numberformat.Format;$f=($fmt -replace '\[[^\]]*\]','' -replace '"[^"]*"','')
 if($f -notmatch '(?i)[dy]' -or $f -match '[#0]'){return ,$Values}
 $inv=[Globalization.CultureInfo]::InvariantCulture;$n=0.0
 for($i=0;$i -lt $Values.Length;$i++){$v=$Values[$i];if($v -ne '' -and [double]::TryParse($v,[Globalization.NumberStyles]::Float,$inv,[ref]$n) -and $n -gt 0 -and $n -lt 2958466){$d=[DateTime]::FromOADate($n);if($d.TimeOfDay.Ticks -eq 0){$Values[$i]=$d.ToString('yyyy-MM-dd')}else{$Values[$i]=$d.ToString('o')}}}
 return ,$Values
}
function Get-CmbTrimmedCount($Cols,[int]$N){
 # Row count without the blank rows at the end (formatted but empty rows that Excel still reports).
 $last=-1;foreach($k in @($Cols.Keys)){$a=$Cols[$k];for($r=$N-1;$r -gt $last;$r--){if($a[$r] -ne ''){$last=$r;break}}};return ($last+1)
}
function Get-CmbColumnStats([string]$Header,[int]$Index,[string[]]$Values,[int]$Rows){
 # One column of a fingerprint: inferred type, fill and distinct counts, shapes, common values (repeating categories only) and examples.
 $counts=New-Object 'Collections.Generic.Dictionary[string,int]';$set=$null;$filled=0
 for($r=0;$r -lt $Rows;$r++){
  $v=$Values[$r];if($v -eq ''){continue};$filled++
  if($null -eq $set){$x=0;if($counts.TryGetValue($v,[ref]$x)){$counts[$v]=$x+1}else{$counts[$v]=1;if($counts.Count -gt 60){$set=[Collections.Generic.HashSet[string]]::new([string[]]@($counts.Keys))}}}
  else{if($set.Count -lt 500000){[void]$set.Add($v)}}
 }
 $distinct=$(if($null -eq $set){$counts.Count}else{$set.Count})
 $kinds=@{};$shapes=@{};$sample=0;$lead=$false;$minL=[int]::MaxValue;$maxL=0;$min=$null;$max=$null;$examples=[Collections.Generic.List[string]]::new()
 $step=[Math]::Max(1,[int][Math]::Floor($filled/1000));$k=0
 for($r=0;$r -lt $Rows;$r++){
  $v=$Values[$r];if($v -eq ''){continue};$k++;if(($k-1)%$step -ne 0){continue};$sample++
  $kind=Get-CmbValueKind $v;$kinds[$kind]=1+[int]$kinds[$kind]
  $sh=Get-CmbShape $v;$shapes[$sh]=1+[int]$shapes[$sh]
  if($v.Length -lt $minL){$minL=$v.Length};if($v.Length -gt $maxL){$maxL=$v.Length};if($v -match '^0\d'){$lead=$true}
  if($examples.Count -lt 3 -and !$examples.Contains($v)){$examples.Add($v)}
  if($kind -eq 'date'){$d=[CmbRt]::Date($v);if($null -eq $min -or $d -lt $min){$min=$d};if($null -eq $max -or $d -gt $max){$max=$d}}
  elseif($kind -in @('integer','number')){$d=[CmbRt]::Num($v);if($null -ne $d){if($null -eq $min -or $d -lt $min){$min=$d};if($null -eq $max -or $d -gt $max){$max=$d}}}
 }
 $type='empty'
 if($sample -gt 0){
  $top=@($kinds.Keys|Sort-Object @{Expression={$kinds[$_]}} -Descending);$t0=[string]$top[0];$share=$kinds[$t0]/$sample
  if($kinds.ContainsKey('integer') -and $kinds.ContainsKey('number') -and ($kinds['integer']+$kinds['number'])/$sample -ge 0.95){$t0='number';$share=1}
  if($share -ge 0.95){$type=$t0}else{$type='mixed ('+(@($top|Select-Object -First 3|ForEach-Object{$_+' '+[Math]::Round(100*$kinds[$_]/$sample)+'%'}) -join ', ')+')'}
  if($type -eq 'integer' -and $lead){$type='text (digits with leading zeros)'}elseif($type -in @('integer','number') -and $null -ne $min -and [double]$min -ge 1000 -and [double]$max -lt 80000 -and $Header -match '(?i)date|\bdob\b|birth|\bdt\b|admit|discharge' -and $Header -notmatch '(?i)count|number|\bnum\b|#|days'){$type='date (stored as Excel day numbers)'}elseif($type -eq 'integer'){$type='number'}
 }
 $pat=@($shapes.Keys|Sort-Object @{Expression={$shapes[$_]}} -Descending|Select-Object -First 3|ForEach-Object{[ordered]@{shape=$_;share=[Math]::Round($shapes[$_]/[Math]::Max(1,$sample),3)}})
 $common=@();if($null -eq $set -and $distinct -gt 0 -and $distinct -le 25 -and $filled -ge 2*$distinct){$common=@($counts.Keys|Sort-Object @{Expression={$counts[$_]}} -Descending|ForEach-Object{[ordered]@{value=$_;rows=$counts[$_]}})}
 $fmt={param($x)if($null -eq $x){return ''};if($x -is [DateTime]){return ([DateTime]$x).ToString('yyyy-MM-dd')};return [CmbRt]::Text($x)}
 return [ordered]@{index=($Index+1);letter=(Get-ColumnLetter ($Index+1));header=$Header;type=$type;filled=$filled;blank=($Rows-$filled);distinct=$distinct;unique=($filled -gt 0 -and $distinct -eq $filled);patterns=$pat;common=$common;examples=@($examples);min=(& $fmt $min);max=(& $fmt $max);minLength=$(if($sample){$minL}else{0});maxLength=$maxL}
}
function Get-CmbFingerprint([string]$Path,[string]$TextName=''){
 $sheets=@(foreach($t in @(Read-CmbSheets $Path '' 0 $null $TextName)){
  $cols=@(for($c=0;$c -lt $t.headers.Count;$c++){$h=$t.headers[$c];Get-CmbColumnStats $h $c ([string[]]$t.columns[$h]) $t.rows})
  $layout=(Get-CmbLayoutId $t.headers)
  $s=[ordered]@{name=$t.name;headerRow=$t.headerRow;rows=$t.rows;layoutId=$layout;headers=@($t.headers);columns=$cols}
  if($t.ContainsKey('delimiter')){$s['delimiter']=$(if($t.delimiter -eq "`t"){'tab'}else{$t.delimiter});$s['encoding']=$t.encoding;$s['malformedLines']=$t.malformed}
  $s
 })
 return @($sheets)
}
function Get-CmbLayoutId([string[]]$Headers){
 # The same report exported on another day has the same column headers in the same order, so the same layout id.
 $sha=[Security.Cryptography.SHA256]::Create();try{$b=$sha.ComputeHash([Text.Encoding]::UTF8.GetBytes((@($Headers|ForEach-Object{ConvertTo-CmbHeaderKey $_}) -join "`n")))}finally{$sha.Dispose()}
 return (([BitConverter]::ToString($b) -replace '-','').Substring(0,10))
}
function Remove-OldCmbFiles{
 $busy=@{};foreach($r in @(Get-CmbRuns)){if([string](Get-P $r 'state' '') -in @('Starting','Running')){foreach($a in @(Get-P $r 'assignments' @())){$busy[[string](Get-P $a 'token' '')]=$true}}}
 foreach($f in @(Get-CmbFiles|Select-Object -Skip $script:CmbFileKeep)){$t=[string](Get-P $f 'token' '');if($t -match '^[a-f0-9]{32}$' -and !$busy.ContainsKey($t)){Remove-CmbFileRecord $f}}
}
function Remove-CmbFileRecord($File){
 $p=[string](Get-P $File 'path' '');if($p -and [IO.Path]::GetFullPath($p).StartsWith([IO.Path]::GetFullPath($script:Paths.CombineFiles),[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue}
 Remove-Item -LiteralPath (Get-CmbFilePath ([string]$File.token)) -Force -ErrorAction SilentlyContinue
}
function Receive-CmbUpload($Context){
 # Saves one chosen export and fingerprints it.
 $name=[Net.WebUtility]::UrlDecode([string]$Context.Request.Headers['X-File-Name']);if([string]::IsNullOrWhiteSpace($name)){throw 'Missing file name.'}
 $name=[IO.Path]::GetFileName($name);$ext=[IO.Path]::GetExtension($name).ToLowerInvariant()
 if($ext -eq '.xls'){throw 'Old .xls workbooks cannot be read; open the file in Excel and save it as .xlsx.'};if($script:CmbExtensions -notcontains $ext){throw ('Only '+($script:CmbExtensions -join ', ')+' files can be loaded.')}
 $token=[Guid]::NewGuid().ToString('N');$path=Join-Path $script:Paths.CombineFiles ($token+$ext)
 $stream=[IO.File]::Create($path);try{$buffer=New-Object byte[] 1048576;$total=0;while(($read=$Context.Request.InputStream.Read($buffer,0,$buffer.Length)) -gt 0){$total+=$read;if($total -gt 2147483648){throw 'Upload exceeds 2 GB limit.'};$stream.Write($buffer,0,$read)}}finally{$stream.Dispose()}
 $fileDate=$null;try{$ms=[string]$Context.Request.Headers['X-File-Timestamp'];if($ms){$fileDate=[DateTimeOffset]::FromUnixTimeMilliseconds([int64]$ms).UtcDateTime.ToString('o')}}catch{}
 $m=Add-CmbFile $path $name $token $fileDate;Remove-OldCmbFiles;return $m
}
function Add-CmbFile([string]$Path,[string]$Name,[string]$Token,$FileDate){
 # Fingerprints a file saved under combine\source-files. A file already loaded (same SHA-256) is reused and marked as loaded now.
 $ext=[IO.Path]::GetExtension($Name).ToLowerInvariant();$sw=[Diagnostics.Stopwatch]::StartNew()
 try{
  $hash=Get-FileHash256 $Path
  foreach($f in @(Get-CmbFiles)){
   if([string](Get-P $f 'sha256' '') -ne $hash -or !(Test-Path -LiteralPath ([string](Get-P $f 'path' '')))){continue}
   Remove-Item -LiteralPath $Path -Force;Set-P $f 'originalName' $Name;Set-P $f 'loadedUtc' ([DateTime]::UtcNow.ToString('o'));if($FileDate){Set-P $f 'fileDateUtc' $FileDate};Save-JsonAtomic (Get-CmbFilePath ([string]$f.token)) $f
   Log 'COMBINE_FILE_LOADED' 'OK' ($Name+' (already loaded)');return (Get-CmbFileModel ([string]$f.token) -Full)
  }
  $sheets=@(Get-CmbFingerprint $Path ([IO.Path]::GetFileNameWithoutExtension($Name)));if($sheets.Count -eq 0){throw 'No table was found in the file (every sheet is empty).'}
  $rec=[ordered]@{cmbFileVersion=1;token=$Token;originalName=$Name;path=$Path;extension=$ext;sha256=$hash;byteLength=(Get-Item -LiteralPath $Path).Length;loadedUtc=[DateTime]::UtcNow.ToString('o');fileDateUtc=$FileDate;seconds=[Math]::Round($sw.Elapsed.TotalSeconds,1);sheets=$sheets}
  Save-JsonAtomic (Get-CmbFilePath $Token) $rec
 }catch{Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue;Log 'COMBINE_FILE' 'FAILED' $_.Exception.Message;throw}
 Log 'COMBINE_FILE_LOADED' 'OK' ($Name+'; '+(@($sheets|ForEach-Object{$_.name+' '+$_.rows+' rows'}) -join ', ')+' in '+$rec.seconds+'s')
 return (Get-CmbFileModel $Token -Full)
}
function Import-CmbCanonicalFiles{
 # The Provider Analysis sources (canonical-current) loaded into Combine under their canonical names; unchanged files are reused.
 $out=@()
 foreach($src in @(Registry)){
  $p=Join-Path $script:Paths.CanonicalCurrent ([string]$src.canonicalFileName);if(!(Test-Path -LiteralPath $p)){continue}
  $tok=[Guid]::NewGuid().ToString('N');$dest=Join-Path $script:Paths.CombineFiles ($tok+'.xlsx');Copy-Item -LiteralPath $p -Destination $dest -Force
  $m=Add-CmbFile $dest ([string]$src.canonicalFileName) $tok ((Get-Item -LiteralPath $p).LastWriteTimeUtc.ToString('o'))
  $out+=[ordered]@{sourceKey=[string]$src.sourceKey;file=[string]$m.originalName;token=[string]$m.token}
 }
 Remove-OldCmbFiles;return @($out)
}
function Get-CmbFileModel([string]$Token,[switch]$Full){
 $f=Get-CmbFile $Token
 $sheets=@(foreach($s in @(Get-P $f 'sheets' @())){
  $m=[ordered]@{name=[string]$s.name;headerRow=[int]$s.headerRow;rows=[int]$s.rows;layoutId=[string]$s.layoutId;columnCount=@(Get-P $s 'columns' @()).Count}
  if($Full){$m['columns']=@(Get-P $s 'columns' @());foreach($k in @('delimiter','encoding','malformedLines')){$v=Get-P $s $k $null;if($null -ne $v){$m[$k]=$v}}}
  $m
 })
 return [ordered]@{token=[string]$f.token;originalName=[string]$f.originalName;extension=[string](Get-P $f 'extension' '');byteLength=[int64](Get-P $f 'byteLength' 0);loadedUtc=(ConvertTo-IsoText (Get-P $f 'loadedUtc' ''));fileDateUtc=(ConvertTo-IsoText (Get-P $f 'fileDateUtc' ''));seconds=(Get-P $f 'seconds' $null);sheets=$sheets}
}

# --- Combine configs: combine\configs\<name>.json, each save keeping the previous version under configs\history\<name>\.
function Get-CmbConfigName([string]$Name){$n=Get-SafeFileName $Name;if(!$n -or $n -eq 'Provider'){$n=($Name -replace '[^A-Za-z0-9 _.-]','').Trim()};if(!$n){throw 'Give the config a name.'};return $n}
function Get-CmbConfigPath([string]$Name){$n=Get-CmbConfigName $Name;return (Join-Path $script:Paths.CombineConfigs ($n+'.json'))}
function Get-CmbConfigs{
 return @(Get-ChildItem $script:Paths.CombineConfigs -Filter '*.json' -File -ErrorAction SilentlyContinue|Sort-Object Name|ForEach-Object{
  $text='';try{$text=[IO.File]::ReadAllText($_.FullName)}catch{}
  $r=Read-CmbConfigText $text;$cfg=@($r.rules|Where-Object{[string](Get-CmbValue $_ 'rule' '') -ieq 'config'})
  $desc='';if($cfg.Count){$desc=Get-CmbText $cfg[0] 'description'}
  [ordered]@{name=$_.BaseName;path=$_.FullName;rules=@($r.rules).Count;readErrors=@($r.errors).Count;description=$desc;modifiedUtc=$_.LastWriteTimeUtc.ToString('o')}
 })
}
function Read-CmbConfigFile([string]$Name){
 $p=Get-CmbConfigPath $Name;if(!(Test-Path -LiteralPath $p)){throw ('There is no config named "'+$Name+'".')}
 $text=[IO.File]::ReadAllText($p);$r=Read-CmbConfigText $text
 return @{name=[IO.Path]::GetFileNameWithoutExtension($p);path=$p;text=$text;rules=@($r.rules);errors=@($r.errors)}
}
function Get-CmbHistoryDir([string]$Name){return (Join-Path $script:Paths.CombineHistory (Get-CmbConfigName $Name))}
function Save-CmbConfig([string]$Name,[object[]]$Rules,[string]$Reason='saved'){
 # Writes the config one rule per line, keeping the previous file in the history folder; the config rule's name follows the file name.
 $n=Get-CmbConfigName $Name;$p=Get-CmbConfigPath $n
 $cfg=@($Rules|Where-Object{[string](Get-CmbValue $_ 'rule' '') -ieq 'config'})
 if($cfg.Count -eq 0){$c=[ordered]@{};$c['rule']='config';$c['name']=$n;$Rules=@(,$c)+@($Rules)}else{$cfg[0]['name']=$n}
 $text=Format-CmbConfigText $Rules
 if(Test-Path -LiteralPath $p){
  $old=[IO.File]::ReadAllText($p);if($old -eq $text){return @{name=$n;path=$p;changed=$false}}
  $dir=Get-CmbHistoryDir $n;if(!(Test-Path $dir)){New-Item -ItemType Directory -Path $dir -Force|Out-Null}
  $stamp=(Get-Item -LiteralPath $p).LastWriteTimeUtc.ToString('yyyyMMdd-HHmmss');$hp=Join-Path $dir ($stamp+'.json');$k=2;while(Test-Path -LiteralPath $hp){$hp=Join-Path $dir ($stamp+'-'+$k+'.json');$k++}
  Copy-Item -LiteralPath $p -Destination $hp
 }
 $tmp=$p+'.'+$PID+'.tmp';[IO.File]::WriteAllText($tmp,$text,(New-Object Text.UTF8Encoding($false)))
 if(Test-Path -LiteralPath $p){[IO.File]::Replace($tmp,$p,[NullString]::Value)}else{[IO.File]::Move($tmp,$p)}
 Log 'COMBINE_CONFIG_SAVED' 'OK' ($n+' ('+$Reason+', '+@($Rules).Count+' rules)')
 return @{name=$n;path=$p;changed=$true}
}
function Get-CmbConfigHistory([string]$Name){
 $dir=Get-CmbHistoryDir $Name;if(!(Test-Path $dir)){return @()}
 return @(Get-ChildItem $dir -Filter '*.json' -File|Sort-Object Name -Descending|ForEach-Object{$r=Read-CmbConfigText ([IO.File]::ReadAllText($_.FullName));[ordered]@{version=$_.BaseName;savedUtc=$_.LastWriteTimeUtc.ToString('o');rules=@($r.rules).Count}})
}
function Read-CmbConfigVersion([string]$Name,[string]$Version){
 if($Version -notmatch '^\d{8}-\d{6}(-\d+|-deleted)?$'){throw 'Invalid version.'};$p=Join-Path (Get-CmbHistoryDir $Name) ($Version+'.json');if(!(Test-Path -LiteralPath $p)){throw 'That version was not found.'}
 return [IO.File]::ReadAllText($p)
}
function Remove-CmbConfig([string]$Name){
 # The file moves into its history folder, so a deleted config can still be recovered by hand.
 $p=Get-CmbConfigPath $Name;if(!(Test-Path -LiteralPath $p)){throw 'That config was not found.'}
 $dir=Get-CmbHistoryDir $Name;if(!(Test-Path $dir)){New-Item -ItemType Directory -Path $dir -Force|Out-Null}
 Move-Item -LiteralPath $p -Destination (Join-Path $dir ((Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmss')+'-deleted.json')) -Force;Log 'COMBINE_CONFIG_DELETED' 'OK' $Name
}
# --- Combine plan: checks a config against the rule reference (and the matched files' headers when known) and organizes it for a run.
function Test-CmbName([string]$Name){return ($Name -match '^[A-Za-z][A-Za-z0-9_]*$')}
function Get-CmbTypeOf([string]$Type){$t=$Type.Trim().ToLowerInvariant();if($t -in @('','text','string')){return 'text'};if($t -in @('number','numeric','integer','decimal')){return 'number'};if($t -in @('date','datetime')){return 'date'};if($t -in @('bool','boolean','yesno','flag')){return 'bool'};return $null}
function ConvertTo-CmbCleanSteps($Steps){
 # Normalized cleaning steps (@{op;n;text;with;map;else}), or a message saying which step is wrong.
 $out=[Collections.Generic.List[object]]::new();$known=@('trim','upper','lower','title','collapsespaces','digits','alnum','stripleadingzeros','removebracketid','namekey','filldown')
 $list=@();if($null -eq $Steps){return @{steps=@();error=''}};if($Steps -is [Array]){$list=$Steps}else{$list=@($Steps)}
 foreach($s in $list){
  if($s -is [string]){
   $x=$s.Trim();$lx=$x.ToLowerInvariant()
   if($known -contains $lx){$out.Add(@{op=$lx});continue}
   if($lx -match '^(pad|left):\s*(\d{1,3})$'){$out.Add(@{op=$matches[1];n=[int]$matches[2]});continue}
   if($lx -match '^blankif:'){$out.Add(@{op='blankif';text=$x.Substring(8).Trim()});continue}
   return @{steps=@();error=('unknown cleaning step "'+$x+'"')}
  }
  if($s -is [Collections.IDictionary]){
   if($s.Contains('replace')){$p=[string]$s['replace'];try{$null=[regex]::new($p)}catch{return @{steps=@();error=('"replace" pattern "'+$p+'" is not a valid regular expression')}};$out.Add(@{op='replace';rx=$p;with=[string](Get-CmbValue $s 'with' '')});continue}
   if($s.Contains('extract')){$p=[string]$s['extract'];try{$null=[regex]::new($p)}catch{return @{steps=@();error=('"extract" pattern "'+$p+'" is not a valid regular expression')}};$out.Add(@{op='extract';rx=$p});continue}
   if($s.Contains('map')){
    $m=$s['map'];if(!($m -is [Collections.IDictionary])){return @{steps=@();error='"map" needs an object of "from":"to" pairs'}}
    $h=@{};foreach($k in @($m.Keys)){$h[[CmbRt]::Key($k)]=$(if($null -eq $m[$k]){$null}else{[string]$m[$k]})}
    $else=[string](Get-CmbValue $s 'else' 'keep');$out.Add(@{op='map';map=$h;else=$else});continue
   }
   return @{steps=@();error=('unknown cleaning step '+(ConvertTo-CmbJson $s))}
  }
  return @{steps=@();error=('unknown cleaning step '+(ConvertTo-CmbJson $s))}
 }
 return @{steps=@($out);error=''}
}
function Get-CmbClosest([string]$Name,[string[]]$Options){
 # The option most like Name (for "did you mean"), by edit distance ignoring case, spaces and punctuation; '' when nothing is close.
 $n=($Name -replace '[^A-Za-z0-9]','').ToLowerInvariant();if(!$n){return ''};$best='';$bestD=[int]::MaxValue
 foreach($o in $Options){
  if($null -eq $o){continue};$x=([string]$o -replace '[^A-Za-z0-9]','').ToLowerInvariant();if(!$x){continue};if($x -eq $n){return [string]$o}
  $prev=New-Object int[] ($x.Length+1);for($j=0;$j -le $x.Length;$j++){$prev[$j]=$j}
  for($i=1;$i -le $n.Length;$i++){$cur=New-Object int[] ($x.Length+1);$cur[0]=$i;for($j=1;$j -le $x.Length;$j++){$c=$(if($n[$i-1] -eq $x[$j-1]){0}else{1});$cur[$j]=[Math]::Min([Math]::Min($prev[$j]+1,$cur[$j-1]+1),$prev[$j-1]+$c)};$prev=$cur}
  $dist=$prev[$x.Length];if($dist -lt $bestD){$bestD=$dist;$best=[string]$o}
 }
 if($bestD -le [Math]::Max(2,[int][Math]::Floor($n.Length/3))){return $best};return ''
}
function Get-CmbPlan([object[]]$Rules,[hashtable]$Tables=@{},[switch]$Strict){
 # Tables: source id (lower case) -> @{headers;file} for sources whose file is known. Strict (a run) makes a missing file an error.
 $P=@{problems=[Collections.Generic.List[object]]::new();config=@{name='';description='';asOf=(Get-Date).Date;asOfText='';sheet='Patients';lineage=$true;booleans='Y/N'};sources=[Collections.Generic.List[object]]::new();sourceByKey=@{};keys=[Collections.Generic.List[object]]::new();fields=[ordered]@{};derives=[Collections.Generic.List[object]]::new();lookups=[ordered]@{};patientFilters=[Collections.Generic.List[object]]::new();sorts=[Collections.Generic.List[object]]::new();outputs=[Collections.Generic.List[object]]::new();views=[ordered]@{};fieldTypes=@{};fieldNames=@{};order=[Collections.Generic.List[string]]::new();rules=@($Rules)}
 $add={param([int]$i,[string]$sev,[string]$msg)$P.problems.Add([ordered]@{rule=$i;severity=$sev;message=$msg})}
 $byType=@{};for($i=0;$i -lt $Rules.Count;$i++){
  $r=$Rules[$i];$no=$i+1
  if(!($r -is [Collections.IDictionary])){& $add $no 'error' 'This row is not a rule object.';continue}
  $type=([string](Get-CmbValue $r 'rule' '')).Trim().ToLowerInvariant()
  if(!$type){& $add $no 'error' 'Missing "rule" (the rule type).';continue}
  if(!$script:CmbRuleSpec.Contains($type)){$g=Get-CmbClosest $type @($script:CmbRuleSpec.Keys);& $add $no 'error' ('Unknown rule type "'+$type+'".'+$(if($g){' Did you mean "'+$g+'"?'}else{''})+' Rule types: '+(@($script:CmbRuleSpec.Keys) -join ', ')+'.');continue}
  $allowed=@($script:CmbRuleSpec[$type].props.Keys)+@('rule','note','id')
  foreach($k in @($r.Keys)){if($allowed -notcontains [string]$k){$g=Get-CmbClosest ([string]$k) $allowed;& $add $no 'warning' ('Unknown property "'+$k+'" is ignored.'+$(if($g){' Did you mean "'+$g+'"?'}else{' A '+$type+' rule takes: '+(@($script:CmbRuleSpec[$type].props.Keys) -join ', ')+'.'}))}}
  if(!$byType.ContainsKey($type)){$byType[$type]=[Collections.Generic.List[int]]::new()};$byType[$type].Add($i)
 }
 $of={param([string]$t)if($byType.ContainsKey($t)){return ,@($byType[$t])};return ,@()}
 # config
 $cfgs=& $of 'config'
 if($cfgs.Count -gt 1){foreach($i in @($cfgs|Select-Object -Skip 1)){& $add ($i+1) 'error' 'Only one config rule is allowed.'}}
 if($cfgs.Count -ge 1){
  $r=$Rules[$cfgs[0]];$no=$cfgs[0]+1;$P.config.name=Get-CmbText $r 'name';$P.config.description=Get-CmbText $r 'description'
  $a=Get-CmbText $r 'asOf';if($a){$d=[CmbRt]::Date($a);if($null -eq $d){& $add $no 'error' ('asOf "'+$a+'" is not a date; write it as YYYY-MM-DD.')}else{$P.config.asOf=([DateTime]$d).Date;$P.config.asOfText=([DateTime]$d).ToString('yyyy-MM-dd')}}
  $s=Get-CmbText $r 'sheet';if($s){$P.config.sheet=$s}
  $l=Get-CmbValue $r 'lineage' $true;if($l -is [bool]){$P.config.lineage=$l}else{& $add $no 'error' 'lineage must be true or false.'}
  $b=Get-CmbText $r 'booleans';if($b){$hit=@(@('Y/N','Yes/No','TRUE/FALSE','1/0')|Where-Object{$_ -ieq $b});if($hit.Count){$P.config.booleans=$hit[0]}else{& $add $no 'error' 'booleans must be "Y/N", "Yes/No", "TRUE/FALSE" or "1/0".'}}
 }else{& $add 0 'warning' 'There is no config rule; one is added (named after the file) when the config is saved.'}
 # sources
 foreach($i in (& $of 'source')){
  $r=$Rules[$i];$no=$i+1;$id=Get-CmbText $r 'source'
  if(!$id){& $add $no 'error' 'A source rule needs "source": a short id such as "epic".';continue}
  if(!(Test-CmbName $id)){& $add $no 'error' ('Source id "'+$id+'" must use letters, digits and underscore and start with a letter.');continue}
  $key=$id.ToLowerInvariant();if($P.sourceByKey.ContainsKey($key)){& $add $no 'error' ('Source "'+$id+'" is defined twice.');continue}
  $match=@(Get-CmbList $r 'match');if($match.Count -eq 0){& $add $no 'error' ('Source "'+$id+'" needs "match": the column headers its file must have.')}
  $role=(Get-CmbText $r 'role').ToLowerInvariant();if(!$role){$role='patients'};if($role -notin @('patients','lookup')){& $add $no 'error' 'role must be "patients" or "lookup".';$role='patients'}
  $hr=Get-CmbValue $r 'headerRow' 0;$hrN=0;if($hr -is [double] -and $hr -ge 1 -and $hr -le 1000 -and [Math]::Floor($hr) -eq $hr){$hrN=[int]$hr}elseif([string]$hr -ne '0'){& $add $no 'error' 'headerRow must be a row number such as 1.'}
  $req=Get-CmbValue $r 'required' $true;if(!($req -is [bool])){& $add $no 'error' 'required must be true or false.';$req=$true}
  $src=@{id=$id;key=$key;ruleIndex=$i;label=$(if((Get-CmbText $r 'label')){Get-CmbText $r 'label'}else{$id});match=$match;fileName=(Get-CmbText $r 'fileName');sheet=(Get-CmbText $r 'sheet');headerRow=$hrN;role=$role;required=$req;lookupKey=(Get-CmbText $r 'key');unique=@(Get-CmbList $r 'unique');read='values';maps=[Collections.Generic.List[object]]::new();filters=[Collections.Generic.List[object]]::new();fields=[ordered]@{};rank=$P.sources.Count;table=$null}
  $rd=(Get-CmbText $r 'read').ToLowerInvariant();if($rd){if($rd -in @('values','text')){$src.read=$rd}else{& $add $no 'error' 'read must be "values" or "text".'}}
  if($Tables.ContainsKey($key)){$src.table=$Tables[$key]}
  elseif($Strict -and $req -and $match.Count){& $add $no 'error' ('No file is chosen for source "'+$id+'".')}
  if($role -eq 'lookup' -and !$src.lookupKey){& $add $no 'error' ('Lookup source "'+$id+'" needs "key": the field that lookup() matches on.')}
  if($role -eq 'patients' -and $src.lookupKey){& $add $no 'warning' '"key" is only used by lookup sources; patients are matched by key rules.'}
  $P.sources.Add($src);$P.sourceByKey[$key]=$src
 }
 if($P.sources.Count -eq 0){& $add 0 'error' 'Add at least one source rule.'}
 $headerOf={param($src)if($null -eq $src.table){return $null};$h=@{};foreach($x in @($src.table.headers)){$h[(ConvertTo-CmbHeaderKey $x)]=$x};return $h}
 foreach($src in $P.sources){
  $h=& $headerOf $src;if($null -eq $h){continue}
  $miss=@($src.match|Where-Object{!$h.ContainsKey((ConvertTo-CmbHeaderKey $_))})
  if($miss.Count){& $add ($src.ruleIndex+1) 'error' ('The file chosen for "'+$src.id+'" ('+$src.table.file+') has no column '+(@($miss|ForEach-Object{'"'+$_+'"'}) -join ', ')+'.')}
 }
 # maps
 foreach($i in (& $of 'map')){
  $r=$Rules[$i];$no=$i+1;$sid=Get-CmbText $r 'source';$field=Get-CmbText $r 'field'
  if(!$sid -or !$P.sourceByKey.ContainsKey($sid.ToLowerInvariant())){& $add $no 'error' $(if($sid){'Map rule names source "'+$sid+'", which has no source rule.'}else{'A map rule needs "source".'});continue}
  $src=$P.sourceByKey[$sid.ToLowerInvariant()]
  if(!$field -or !(Test-CmbName $field)){& $add $no 'error' $(if($field){'Field name "'+$field+'" must use letters, digits and underscore and start with a letter.'}else{'A map rule needs "field".'});continue}
  $fk=$field.ToLowerInvariant();if($src.fields.Contains($fk)){& $add $no 'error' ('Source "'+$src.id+'" maps field "'+$field+'" twice.');continue}
  $modes=@(foreach($m in @('column','columns','value','expr','special')){if($r.Contains($m) -and $null -ne $r[$m]){$m}})
  if($modes.Count -ne 1){& $add $no 'error' $(if($modes.Count){'Use only one of '+($modes -join ', ')+'.'}else{'A map rule needs one of column, columns, value, expr or special.'});continue}
  $type='text';$tt=Get-CmbText $r 'type';if($tt){$type=Get-CmbTypeOf $tt;if(!$type){& $add $no 'error' ('type "'+$tt+'" is not one of text, number, date, bool.');$type='text'}}
  $cl=ConvertTo-CmbCleanSteps (Get-CmbValue $r 'clean' $null);if($cl.error){& $add $no 'error' ('clean: '+$cl.error+'. See the cleaning steps in the reference.')}
  $m=@{ruleIndex=$i;field=$field;key=$fk;mode=$modes[0];type=$type;clean=$cl.steps;column='';columns=@();sep=' ';value=$null;expr='';special=''}
  $h=& $headerOf $src
  switch($modes[0]){
   'column'{$m.column=Get-CmbText $r 'column';if(!$m.column){& $add $no 'error' 'column is empty.'}elseif($null -ne $h -and !$h.ContainsKey((ConvertTo-CmbHeaderKey $m.column))){$g=Get-CmbClosest $m.column @($src.table.headers);& $add $no 'error' ('Column "'+$m.column+'" is not in '+$src.table.file+'.'+$(if($g){' Did you mean "'+$g+'"?'}else{''}))}}
   'columns'{$m.columns=@(Get-CmbList $r 'columns');$m.sep=[string](Get-CmbValue $r 'sep' ' ');if($m.columns.Count -eq 0){& $add $no 'error' 'columns is empty.'};if($null -ne $h){foreach($c in $m.columns){if(!$h.ContainsKey((ConvertTo-CmbHeaderKey $c))){$g=Get-CmbClosest $c @($src.table.headers);& $add $no 'error' ('Column "'+$c+'" is not in '+$src.table.file+'.'+$(if($g){' Did you mean "'+$g+'"?'}else{''}))}}}}
   'value'{$m.value=Get-CmbValue $r 'value' $null;if($m.value -is [Collections.IDictionary] -or $m.value -is [Array]){& $add $no 'error' 'value must be text, a number, true or false.'}}
   'expr'{$m.expr=Get-CmbText $r 'expr'}
   'special'{$m.special=(Get-CmbText $r 'special');$hit=@(@('fileName','fileDate','rowNumber','sourceLabel')|Where-Object{$_ -ieq $m.special});if($hit.Count){$m.special=$hit[0]}else{& $add $no 'error' 'special must be "fileName", "fileDate", "rowNumber" or "sourceLabel".'}}
  }
  if($m.mode -eq 'expr'){
   $names=@{};if($null -ne $h){foreach($k in @($h.Keys)){$names[$k]=0}};foreach($k in @($src.fields.Keys)){$names[$k]=0}
   try{$null=New-CmbCompiled $m.expr (New-CmbExprContext $names (Get-CmbLookupTablesForCheck $P $Rules) $P.config.asOf ('not a column of '+$(if($src.table){$src.table.file}else{'the file'})+' or a field mapped above') -AcceptAny:($null -eq $h))}catch{& $add $no 'error' $_.Exception.Message}
  }
  if($m.special -eq 'fileDate' -and $type -eq 'text'){$m.type='date'};if($m.special -eq 'rowNumber' -and $type -eq 'text'){$m.type='number'}
  $src.maps.Add($m);$src.fields[$fk]=@{name=$field;type=$m.type;map=$m}
 }
 foreach($src in $P.sources){
  if($src.maps.Count -eq 0){& $add ($src.ruleIndex+1) 'warning' ('Source "'+$src.id+'" has no map rules, so nothing is read from it.')}
  if($src.role -eq 'lookup' -and $src.lookupKey -and !$src.fields.Contains($src.lookupKey.ToLowerInvariant())){& $add ($src.ruleIndex+1) 'error' ('Lookup key "'+$src.lookupKey+'" is not a field mapped from "'+$src.id+'".')}
  $badU=@($src.unique|Where-Object{!$src.fields.Contains($_.ToLowerInvariant())});if($badU.Count){& $add ($src.ruleIndex+1) 'error' ('unique: '+(@($badU|ForEach-Object{'"'+$_+'"'}) -join ', ')+' is not a field mapped from "'+$src.id+'".')}
 }
 # Field names and their types across patient sources (first map wins the display name and type).
 foreach($src in $P.sources){if($src.role -ne 'patients'){continue};foreach($k in @($src.fields.Keys)){if(!$P.fieldNames.ContainsKey($k)){$P.fieldNames[$k]=$src.fields[$k].name;$P.fieldTypes[$k]=$src.fields[$k].type}elseif($P.fieldTypes[$k] -ne $src.fields[$k].type){& $add ($src.fields[$k].map.ruleIndex+1) 'warning' ('Field "'+$src.fields[$k].name+'" is '+$src.fields[$k].type+' here but '+$P.fieldTypes[$k]+' in another source; values compare as '+$P.fieldTypes[$k]+'.')}}}
 # source filters
 foreach($i in (& $of 'filter')){
  $r=$Rules[$i];$no=$i+1;$sid=Get-CmbText $r 'source';$keep=Get-CmbText $r 'keep';$drop=Get-CmbText $r 'drop'
  if(($keep -and $drop) -or (!$keep -and !$drop)){& $add $no 'error' 'A filter rule needs keep or drop (one of them).';continue}
  $f=@{ruleIndex=$i;keep=[bool]$keep;expr=$(if($keep){$keep}else{$drop})}
  if(!$sid){$P.patientFilters.Add($f);continue}
  if(!$P.sourceByKey.ContainsKey($sid.ToLowerInvariant())){& $add $no 'error' ('Filter names source "'+$sid+'", which has no source rule.');continue}
  $src=$P.sourceByKey[$sid.ToLowerInvariant()];$h=& $headerOf $src;$names=@{};if($null -ne $h){foreach($k in @($h.Keys)){$names[$k]=0}};foreach($k in @($src.fields.Keys)){$names[$k]=0}
  try{$null=New-CmbCompiled $f.expr (New-CmbExprContext $names (Get-CmbLookupTablesForCheck $P $Rules) $P.config.asOf ('not a column of '+$(if($src.table){$src.table.file}else{'the file'})+' or a field of source "'+$src.id+'"') -AcceptAny:($null -eq $h)) -AsBool}catch{& $add $no 'error' $_.Exception.Message}
  $src.filters.Add($f)
 }
 # keys
 $patientSources=@($P.sources|Where-Object{$_.role -eq 'patients'})
 foreach($i in (& $of 'key')){
  $r=$Rules[$i];$no=$i+1;$fields=@(Get-CmbList $r 'fields');if($fields.Count -eq 0){& $add $no 'error' 'A key rule needs "fields", e.g. ["MRN"].';continue}
  $bad=@($fields|Where-Object{!$P.fieldNames.ContainsKey($_.ToLowerInvariant())});if($bad.Count){& $add $no 'error' ('Key field '+(@($bad|ForEach-Object{'"'+$_+'"'}) -join ', ')+' is not mapped by any patient source.');continue}
  $want=@(Get-CmbList $r 'sources');$use=@()
  if($want.Count){foreach($w in $want){if(!$P.sourceByKey.ContainsKey($w.ToLowerInvariant())){& $add $no 'error' ('Key rule names source "'+$w+'", which has no source rule.')}else{$s=$P.sourceByKey[$w.ToLowerInvariant()];$miss=@($fields|Where-Object{!$s.fields.Contains($_.ToLowerInvariant())});if($miss.Count){& $add $no 'error' ('Source "'+$s.id+'" does not map '+(@($miss|ForEach-Object{'"'+$_+'"'}) -join ', ')+'.')}else{$use+=$s.id}}}}
  else{$use=@(foreach($s in $patientSources){$all=$true;foreach($f in $fields){if(!$s.fields.Contains($f.ToLowerInvariant())){$all=$false}};if($all){$s.id}})}
  $mt=(Get-CmbText $r 'match').ToLowerInvariant();if(!$mt){$mt='exact'};if($mt -notin @('exact','prefix')){& $add $no 'error' 'match must be "exact" or "prefix".';$mt='exact'}
  $ag=@();$agWant=@(Get-CmbList $r 'against')
  if($mt -eq 'prefix'){
   if($agWant.Count){foreach($w in $agWant){if(!$P.sourceByKey.ContainsKey($w.ToLowerInvariant())){& $add $no 'error' ('against names source "'+$w+'", which has no source rule.')}else{$s=$P.sourceByKey[$w.ToLowerInvariant()];$miss=@($fields|Where-Object{!$s.fields.Contains($_.ToLowerInvariant())});if($miss.Count){& $add $no 'error' ('Source "'+$s.id+'" does not map '+(@($miss|ForEach-Object{'"'+$_+'"'}) -join ', ')+'.')}elseif($use -contains $s.id){& $add $no 'error' ('Source "'+$s.id+'" cannot be in both sources and against.')}else{$ag+=$s.id}}}}
   else{$ag=@(foreach($s in $patientSources){if($use -contains $s.id){continue};$all=$true;foreach($f in $fields){if(!$s.fields.Contains($f.ToLowerInvariant())){$all=$false}};if($all){$s.id}})}
   if($ag.Count -eq 0){& $add $no 'error' 'A prefix key needs other sources to match against ("against").'}
  }elseif($agWant.Count){& $add $no 'warning' '"against" is only used with match "prefix".'}
  $P.keys.Add(@{ruleIndex=$i;fields=@($fields|ForEach-Object{$P.fieldNames[$_.ToLowerInvariant()]});sources=@($use);match=$mt;against=@($ag);label=(@($fields|ForEach-Object{$P.fieldNames[$_.ToLowerInvariant()]}) -join ' + ')+$(if($mt -eq 'prefix'){' (prefix)'}else{''})})
 }
 if($P.keys.Count -eq 0 -and $patientSources.Count){& $add 0 'error' 'Add a key rule (for example {"rule":"key","fields":["MRN"]}) so rows can be matched to patients.'}
 foreach($s in $patientSources){$hit=$false;foreach($k in $P.keys){if($k.sources -contains $s.id){$hit=$true}};if(!$hit -and $P.keys.Count){& $add ($s.ruleIndex+1) 'warning' ('No key rule applies to source "'+$s.id+'" (it does not map the key fields), so all its rows will be left out.')}}
 # combined fields: explicit field rules, then every other mapped field with how=priority
 $explicit=@{}
 foreach($i in (& $of 'field')){
  $r=$Rules[$i];$no=$i+1;$name=Get-CmbText $r 'field'
  if(!$name -or !(Test-CmbName $name)){& $add $no 'error' $(if($name){'Field name "'+$name+'" must use letters, digits and underscore and start with a letter.'}else{'A field rule needs "field".'});continue}
  $fk=$name.ToLowerInvariant();if($explicit.ContainsKey($fk)){& $add $no 'error' ('Field "'+$name+'" has two field rules.');continue}
  $how=Get-CmbText $r 'how';if(!$how){$how='priority'};$hit=@($script:CmbHows|Where-Object{$_ -ieq $how});if(!$hit.Count){& $add $no 'error' ('how "'+$how+'" is not one of '+($script:CmbHows -join ', ')+'.');continue};$how=$hit[0]
  $ofName=Get-CmbText $r 'of';if(!$ofName){$ofName=$name};$ok=$ofName.ToLowerInvariant()
  if($how -ne 'exists' -and !$P.fieldNames.ContainsKey($ok)){& $add $no 'error' $(if($ofName -ieq $name){'Field "'+$name+'" is not mapped by any patient source; map it, or name the mapped field with "of".'}else{'"of": field "'+$ofName+'" is not mapped by any patient source.'});continue}
  $from=@(Get-CmbList $r 'from');$fromIds=@()
  foreach($w in $from){if(!$P.sourceByKey.ContainsKey($w.ToLowerInvariant())){& $add $no 'error' ('from names source "'+$w+'", which has no source rule.')}else{$s=$P.sourceByKey[$w.ToLowerInvariant()];if($s.role -ne 'patients'){& $add $no 'error' ('from: "'+$s.id+'" is a lookup source.')}elseif($how -ne 'exists' -and !$s.fields.Contains($ok)){& $add $no 'warning' ('Source "'+$s.id+'" does not map "'+$ofName+'", so it adds nothing to this field.')}else{$fromIds+=$s.id}}}
  if($from.Count -eq 0){$fromIds=@(foreach($s in $patientSources){if($how -eq 'exists' -or $s.fields.Contains($ok)){$s.id}})}
  if($how -eq 'exists' -and $from.Count -eq 0){& $add $no 'warning' 'how "exists" without "from" is true for every patient; name the sources in "from".'}
  $by=Get-CmbText $r 'by'
  if($how -in @('latest','earliest')){if(!$by){& $add $no 'error' ('how "'+$how+'" needs "by": the date field that orders the rows.')}elseif(!$P.fieldNames.ContainsKey($by.ToLowerInvariant())){& $add $no 'error' ('by: field "'+$by+'" is not mapped by any patient source.')}else{foreach($sid in $fromIds){$s=$P.sourceByKey[$sid.ToLowerInvariant()];if(!$s.fields.Contains($by.ToLowerInvariant())){& $add $no 'warning' ('Source "'+$s.id+'" does not map "'+$by+'"; its rows count as oldest.')}}}}
  elseif($by){& $add $no 'warning' ('"by" is only used with how latest or earliest.')}
  $conf=(Get-CmbText $r 'conflicts').ToLowerInvariant();if(!$conf){$conf='report'};if($conf -notin @('report','ignore')){& $add $no 'error' 'conflicts must be "report" or "ignore".';$conf='report'}
  $type=$null;$tt=Get-CmbText $r 'type';if($tt){$type=Get-CmbTypeOf $tt;if(!$type){& $add $no 'error' ('type "'+$tt+'" is not one of text, number, date, bool.')}}
  if(!$type){$type=$(if($how -in @('sum','avg','count','countDistinct')){'number'}elseif($how -in @('any','all','exists')){'bool'}elseif($how -eq 'list'){'text'}elseif($P.fieldTypes.ContainsKey($ok)){$P.fieldTypes[$ok]}else{'text'})}
  if($P.fieldNames.ContainsKey($fk) -and $fk -ne $ok){& $add $no 'warning' ('Field "'+$name+'" is also a mapped field; this rule replaces the combined value of the mapped "'+$P.fieldNames[$fk]+'".')}
  $explicit[$fk]=@{name=$name;key=$fk;ruleIndex=$i;of=$P.fieldNames[$ok];ofKey=$ok;how=$how;from=@($fromIds);by=$by;byKey=$by.ToLowerInvariant();sep=[string](Get-CmbValue $r 'sep' '; ');conflicts=$conf;type=$type}
 }
 foreach($k in @($P.fieldNames.Keys)){
  # Mapped fields in the order first mapped; a field rule of the same name takes its place.
  if($explicit.ContainsKey($k)){continue}
  $P.fields[$k]=@{name=$P.fieldNames[$k];key=$k;ruleIndex=-1;of=$P.fieldNames[$k];ofKey=$k;how='priority';from=@(foreach($s in $patientSources){if($s.fields.Contains($k)){$s.id}});by='';byKey='';sep='; ';conflicts='report';type=$P.fieldTypes[$k]}
 }
 $ordered=[ordered]@{};foreach($k in @($P.fieldNames.Keys)){if($explicit.ContainsKey($k)){$ordered[$k]=$explicit[$k]}elseif($P.fields.Contains($k)){$ordered[$k]=$P.fields[$k]}}
 foreach($i in (& $of 'field')){foreach($k in @($explicit.Keys)){if($explicit[$k].ruleIndex -eq $i -and !$ordered.Contains($k)){$ordered[$k]=$explicit[$k]}}}
 $P.fields=$ordered;foreach($k in @($P.fields.Keys)){$P.order.Add($P.fields[$k].name);$P.fieldTypes[$k]=$P.fields[$k].type}
 # lookups (inline tables)
 foreach($i in (& $of 'lookup')){
  $r=$Rules[$i];$no=$i+1;$name=Get-CmbText $r 'table';if(!$name){& $add $no 'error' 'A lookup rule needs "table".';continue}
  $tk=$name.ToLowerInvariant();if($P.lookups.Contains($tk) -or $P.sourceByKey.ContainsKey($tk)){& $add $no 'error' ('There is already a table or source named "'+$name+'".');continue}
  $vals=Get-CmbValue $r 'values' $null;if(!($vals -is [Collections.IDictionary])){& $add $no 'error' '"values" must be an object of "input":"output" pairs.';continue}
  $h=@{};foreach($k in @($vals.Keys)){$v=$vals[$k];if($v -is [Collections.IDictionary] -or $v -is [Array]){& $add $no 'error' ('values: "'+$k+'" must map to text or a number.');continue};$h[[CmbRt]::Key($k)]=[CmbRt]::Convert($v,'')}
  $d=Get-CmbValue $r 'default' $null;if($null -ne $d){$h['*default*']=[CmbRt]::Convert($d,'')}
  $P.lookups[$tk]=@{name=$name;ruleIndex=$i;table=$h}
 }
 # derives, in order; each may use combined fields, earlier derives and the built-ins
 $patientKnown=@{};foreach($k in @($P.fields.Keys)){$patientKnown[$k]=$true};foreach($b in $script:CmbBuiltins){$patientKnown[$b]=$true}
 $tablesCheck=Get-CmbLookupTablesForCheck $P $Rules
 foreach($i in (& $of 'derive')){
  $r=$Rules[$i];$no=$i+1;$name=Get-CmbText $r 'field';$expr=Get-CmbText $r 'expr'
  if(!$name -or !(Test-CmbName $name)){& $add $no 'error' $(if($name){'Field name "'+$name+'" must use letters, digits and underscore and start with a letter.'}else{'A derive rule needs "field".'});continue}
  if(!$expr){& $add $no 'error' 'A derive rule needs "expr".';continue}
  $fk=$name.ToLowerInvariant();if($patientKnown.ContainsKey($fk)){& $add $no 'error' ('"'+$name+'" is already a field; give the derived field another name.');continue}
  $type=$null;$tt=Get-CmbText $r 'type';if($tt){$type=Get-CmbTypeOf $tt;if(!$type){& $add $no 'error' ('type "'+$tt+'" is not one of text, number, date, bool.')}}
  $names=@{};foreach($k in @($patientKnown.Keys)){$names[$k]=0}
  $ok=$true;try{$null=New-CmbCompiled $expr (New-CmbExprContext $names $tablesCheck $P.config.asOf 'not a field defined above')}catch{$ok=$false;& $add $no 'error' $_.Exception.Message}
  $P.derives.Add(@{ruleIndex=$i;field=$name;key=$fk;expr=$expr;type=$type;ok=$ok});$patientKnown[$fk]=$true;$P.order.Add($name);$P.fieldTypes[$fk]=$(if($type){$type}else{''})
 }
 $nameOf=@{};foreach($k in @($P.fields.Keys)){$nameOf[$k]=$P.fields[$k].name};foreach($d in $P.derives){$nameOf[$d.key]=$d.field};foreach($b in $script:CmbBuiltins){$nameOf[$b]=$b}
 $P.nameOf=$nameOf
 $patientNames=@{};foreach($k in @($nameOf.Keys)){$patientNames[$k]=0}
 foreach($f in $P.patientFilters){try{$null=New-CmbCompiled $f.expr (New-CmbExprContext $patientNames $tablesCheck $P.config.asOf 'not a field') -AsBool}catch{& $add ($f.ruleIndex+1) 'error' $_.Exception.Message}}
 # sort, output
 foreach($i in (& $of 'sort')){$r=$Rules[$i];$f=Get-CmbText $r 'field';if(!$f -or !$nameOf.ContainsKey($f.ToLowerInvariant())){& $add ($i+1) 'error' $(if($f){'Sort field "'+$f+'" is not a field.'}else{'A sort rule needs "field".'});continue};$d=Get-CmbValue $r 'desc' $false;if(!($d -is [bool])){& $add ($i+1) 'error' 'desc must be true or false.';$d=$false};$P.sorts.Add(@{ruleIndex=$i;field=$nameOf[$f.ToLowerInvariant()];key=$f.ToLowerInvariant();desc=$d})}
 $seenOut=@{}
 foreach($i in (& $of 'output')){
  $r=$Rules[$i];$f=Get-CmbText $r 'field';if(!$f -or !$nameOf.ContainsKey($f.ToLowerInvariant())){$g='';if($f){$g=Get-CmbClosest $f @($nameOf.Values)};& $add ($i+1) 'error' $(if($f){'Output field "'+$f+'" is not a field.'+$(if($g){' Did you mean "'+$g+'"?'}else{''})}else{'An output rule needs "field".'});continue}
  $fk=$f.ToLowerInvariant();if($seenOut.ContainsKey($fk)){& $add ($i+1) 'warning' ('Field "'+$f+'" is output twice.')};$seenOut[$fk]=$true
  $w=Get-CmbValue $r 'width' $null;if($null -ne $w -and !($w -is [double] -and $w -gt 0 -and $w -le 255)){& $add ($i+1) 'error' 'width must be a number from 1 to 255.';$w=$null}
  $hd=Get-CmbText $r 'header';if(!$hd){$hd=$nameOf[$fk]}
  $P.outputs.Add(@{ruleIndex=$i;field=$nameOf[$fk];key=$fk;header=$hd;format=(Get-CmbText $r 'format');width=$w})
 }
 # views and measures
 foreach($i in (& $of 'view')){
  $r=$Rules[$i];$no=$i+1;$name=Get-CmbText $r 'view';if(!$name){& $add $no 'error' 'A view rule needs "view": its name.';continue}
  $vk=$name.ToLowerInvariant();if($P.views.Contains($vk)){& $add $no 'error' ('View "'+$name+'" is defined twice.');continue}
  if($name.Length -gt 31 -or $name -match '[\[\]:*?/\\]'){& $add $no 'warning' 'Sheet names are at most 31 characters without [ ] : * ? / \; the sheet name will be shortened.'}
  $gb=@(Get-CmbList $r 'groupBy');if($gb.Count -eq 0){& $add $no 'error' 'A view needs "groupBy": the fields to group patients by.'}
  $badG=@($gb|Where-Object{!$nameOf.ContainsKey($_.ToLowerInvariant())});if($badG.Count){& $add $no 'error' ('groupBy field '+(@($badG|ForEach-Object{'"'+$_+'"'}) -join ', ')+' is not a field.')}
  $where=Get-CmbText $r 'where';if($where){try{$null=New-CmbCompiled $where (New-CmbExprContext $patientNames $tablesCheck $P.config.asOf 'not a field') -AsBool}catch{& $add $no 'error' $_.Exception.Message}}
  $tot=Get-CmbValue $r 'total' $true;if(!($tot -is [bool])){& $add $no 'error' 'total must be true or false.';$tot=$true}
  $desc=Get-CmbValue $r 'desc' $false;if(!($desc -is [bool])){& $add $no 'error' 'desc must be true or false.';$desc=$false}
  $P.views[$vk]=@{name=$name;ruleIndex=$i;groupBy=@($gb|Where-Object{$nameOf.ContainsKey($_.ToLowerInvariant())}|ForEach-Object{$nameOf[$_.ToLowerInvariant()]});where=$where;total=$tot;sortBy=(Get-CmbText $r 'sortBy');desc=$desc;measures=[Collections.Generic.List[object]]::new()}
 }
 foreach($i in (& $of 'measure')){
  $r=$Rules[$i];$no=$i+1;$vn=Get-CmbText $r 'view';$name=Get-CmbText $r 'name'
  if(!$vn -or !$P.views.Contains($vn.ToLowerInvariant())){& $add $no 'error' $(if($vn){'Measure names view "'+$vn+'", which has no view rule.'}else{'A measure rule needs "view".'});continue}
  $v=$P.views[$vn.ToLowerInvariant()];if(!$name){& $add $no 'error' 'A measure rule needs "name".';continue}
  if(@($v.measures|Where-Object{$_.name -ieq $name}).Count -or @($v.groupBy|Where-Object{$_ -ieq $name}).Count){& $add $no 'error' ('View "'+$v.name+'" already has a column "'+$name+'".');continue}
  $agg=Get-CmbText $r 'agg';if(!$agg){$agg='count'};$hit=@($script:CmbAggs|Where-Object{$_ -ieq $agg});if(!$hit.Count){& $add $no 'error' ('agg "'+$agg+'" is not one of '+($script:CmbAggs -join ', ')+'.');continue};$agg=$hit[0]
  $field=Get-CmbText $r 'field';if($agg -in @('sum','avg','min','max','countDistinct')){if(!$field){& $add $no 'error' ('agg "'+$agg+'" needs "field".');continue};if(!$nameOf.ContainsKey($field.ToLowerInvariant())){& $add $no 'error' ('Field "'+$field+'" is not a field.');continue};$field=$nameOf[$field.ToLowerInvariant()]}
  $where=Get-CmbText $r 'where';if($where){try{$null=New-CmbCompiled $where (New-CmbExprContext $patientNames $tablesCheck $P.config.asOf 'not a field') -AsBool}catch{& $add $no 'error' $_.Exception.Message}}
  $v.measures.Add(@{ruleIndex=$i;name=$name;agg=$agg;field=$field;where=$where;of=(Get-CmbText $r 'of');per=(Get-CmbText $r 'per');format=(Get-CmbText $r 'format')})
 }
 foreach($vk in @($P.views.Keys)){
  $v=$P.views[$vk]
  foreach($m in $v.measures){if($m.agg -eq 'percent'){foreach($x in @('of','per')){$t=[string]$m[$x];if(!$t -or !@($v.measures|Where-Object{$_.name -ieq $t -and $_.agg -ne 'percent'}).Count){& $add ($m.ruleIndex+1) 'error' ('percent needs "'+$x+'": the name of another (non-percent) measure in view "'+$v.name+'".')}}}}
  if($v.measures.Count -eq 0){$v.measures.Add(@{ruleIndex=-1;name='Patients';agg='count';field='';where='';of='';per='';format=''})}
  if($v.sortBy -and !@($v.measures|Where-Object{$_.name -ieq $v.sortBy}).Count -and !@($v.groupBy|Where-Object{$_ -ieq $v.sortBy}).Count){& $add ($v.ruleIndex+1) 'error' ('sortBy "'+$v.sortBy+'" is not a measure or group field of this view.')}
 }
 $P.errorCount=@($P.problems|Where-Object{$_.severity -eq 'error'}).Count
 return $P
}
function Get-CmbLookupTablesForCheck($Plan,[object[]]$Rules){
 # Table names lookup() may use while checking: inline lookup rules (all of them, in any order) and lookup sources with their fields.
 $t=@{};$n=0
 foreach($r in $Rules){if($r -is [Collections.IDictionary] -and ([string](Get-CmbValue $r 'rule' '')).Trim() -ieq 'lookup'){$name=Get-CmbText $r 'table';if($name){$t[$name.ToLowerInvariant()]=@{index=$n;fields=@()};$n++}}}
 foreach($s in $Plan.sources){if($s.role -eq 'lookup'){$t[$s.key]=@{index=$n;fields=@(foreach($k in @($s.fields.Keys)){$s.fields[$k].name})};$n++}}
 foreach($k in @(Get-CmbProfileTableNames)){if(!$t.ContainsKey($k)){$t[$k]=@{index=$n;fields=@('npi','provider','location','riskPool')}};$n++}
 return $t
}
function Get-CmbProfileTableNames{return @(@(Registry|ForEach-Object{'profile:'+([string]$_.sourceKey).ToLowerInvariant()})+@('profiles'))}
function Get-CmbProfileTables{
 # Built-in lookup tables from the saved Provider Analysis profiles: profile:<source key> finds a provider by the name its rows carry in that
 # source (the HR sources follow the reports' one-of-two rule), and profiles finds one by NPI. Fields: npi, provider, location, riskPool.
 $t=[ordered]@{};$keys=@(Registry|ForEach-Object{[string]$_.sourceKey});foreach($k in $keys){$t['profile:'+$k.ToLowerInvariant()]=@{}};$byNpi=@{}
 foreach($p in @(Get-ProviderProfiles)){
  if($null -eq $p){continue};$npi=[string](Get-P $p 'npi' '');if($npi -notmatch '^\d{10}$'){continue}
  $row=@{npi=$npi;provider=[string](Get-P $p 'displayName' $npi);location=([string](Get-P $p 'location' '')).Trim();riskPool=([string](Get-P $p 'riskPool' '')).Trim()}
  if(!$byNpi.ContainsKey($npi)){$byNpi[$npi]=$row}
  $a=Get-P $p 'aliases' $null
  foreach($k in $keys){$al=Get-OverviewAlias $a $k;if(!$al){continue};$kk=[CmbRt]::Key($al);$tb=$t['profile:'+$k.ToLowerInvariant()];if(!$tb.ContainsKey($kk)){$tb[$kk]=$row}}
 }
 $t['profiles']=$byNpi;return $t
}
function Get-CmbRuleSummary($Rule){
 # One plain sentence per rule for the rules table and the audit sheet.
 if(!($Rule -is [Collections.IDictionary])){return '(not a rule object)'}
 $t=([string](Get-CmbValue $Rule 'rule' '')).Trim().ToLowerInvariant();$v={param($k)Get-CmbText $Rule $k};$l={param($k),@(Get-CmbList $Rule $k)}
 $q={param([string[]]$x)(@($x|ForEach-Object{'"'+$_+'"'}) -join ', ')}
 switch($t){
  'config'{$s='Config "'+(& $v 'name')+'"';if(& $v 'asOf'){$s+='; today() is '+(& $v 'asOf')};if(& $v 'sheet'){$s+='; patient sheet "'+(& $v 'sheet')+'"'};$lin=Get-CmbValue $Rule 'lineage' $true;if($lin -eq $false){$s+='; no value-source sheet'};return $s}
  'source'{$s='Source '+(& $v 'source');if((& $v 'label')){$s+=' ('+(& $v 'label')+')'};if((& $v 'role') -ieq 'lookup'){$s+=': lookup table keyed on '+(& $v 'key')};$s+=': a file with columns '+(& $q (& $l 'match'));if(& $v 'fileName'){$s+=', named like '+(& $v 'fileName')};if(& $v 'sheet'){$s+=', sheet '+(& $v 'sheet')};if((Get-CmbValue $Rule 'required' $true) -eq $false){$s+=' (optional)'};return $s}
  'map'{
   $s=(& $v 'source')+'.'+(& $v 'field')+' from '
   if($Rule.Contains('column')){$s+='column "'+(& $v 'column')+'"'}elseif($Rule.Contains('columns')){$s+='columns '+(& $q (& $l 'columns'))+' joined'}elseif($Rule.Contains('value')){$s+='the value '+(ConvertTo-CmbJson (Get-CmbValue $Rule 'value' ''))}elseif($Rule.Contains('expr')){$s+=(& $v 'expr')}elseif($Rule.Contains('special')){$s+='the '+(& $v 'special')}else{$s+='(nothing)'}
   $c=Get-CmbValue $Rule 'clean' $null;if($null -ne $c){$s+='; cleaned: '+(@(@($c)|ForEach-Object{if($_ -is [string]){$_}else{ConvertTo-CmbJson $_}}) -join ', ')}
   $ty=& $v 'type';if($ty -and $ty -ine 'text'){$s+='; as '+$ty};return $s
  }
  'filter'{$who=$(if(& $v 'source'){(& $v 'source')+' rows'}else{'Patients'});if(& $v 'keep'){return ($who+': keep where '+(& $v 'keep'))};return ($who+': drop where '+(& $v 'drop'))}
  'key'{$s='Same patient when '+((& $l 'fields') -join ' + ')+' match';$src=& $l 'sources';if($src.Count){$s+=' (sources '+($src -join ', ')+')'};return $s}
  'field'{
   $f=& $v 'field';$how=& $v 'how';if(!$how){$how='priority'};$of=& $v 'of';$src=& $l 'from';$fromText=$(if($src.Count){' from '+($src -join ', ')}else{''})
   $what=$(if($of -and $of -ine $f){$of}else{'value'})
   switch($how.ToLowerInvariant()){
    'priority'{return ($f+' = first non-blank '+$what+$(if($src.Count){' trying '+($src -join ', ')+' in order'}else{' in source order'}))}
    'latest'{return ($f+' = '+$what+' from the row with the latest '+(& $v 'by')+$fromText)}
    'earliest'{return ($f+' = '+$what+' from the row with the earliest '+(& $v 'by')+$fromText)}
    'exists'{return ($f+' = true when the patient is in '+$(if($src.Count){$src -join ' or '}else{'any source'}))}
    'list'{return ($f+' = every different '+$what+$fromText+', joined')}
    'mostcommon'{return ($f+' = most common '+$what+$fromText)}
    'countdistinct'{return ($f+' = number of different '+$what+'s'+$fromText)}
    default{return ($f+' = '+$how+' of '+$(if($of){$of}else{$f})+$fromText)}
   }
  }
  'derive'{return ((& $v 'field')+' = '+(& $v 'expr'))}
  'lookup'{$vals=Get-CmbValue $Rule 'values' $null;$n=0;if($vals -is [Collections.IDictionary]){$n=$vals.Count};return ('Table '+(& $v 'table')+': '+$n+' value'+$(if($n -eq 1){''}else{'s'})+$(if($Rule.Contains('default')){', otherwise '+(ConvertTo-CmbJson $Rule['default'])}else{''}))}
  'sort'{return ('Sort by '+(& $v 'field')+$(if((Get-CmbValue $Rule 'desc' $false) -eq $true){', largest or latest first'}else{''}))}
  'output'{$h=& $v 'header';$f=& $v 'field';return ('Column "'+$(if($h){$h}else{$f})+'"'+$(if($h -and $h -cne $f){' = '+$f}else{''})+$(if(& $v 'format'){' as '+(& $v 'format')}else{''}))}
  'view'{$s='View "'+(& $v 'view')+'": patients grouped by '+((& $l 'groupBy') -join ', ');if(& $v 'where'){$s+=' where '+(& $v 'where')};return $s}
  'measure'{$agg=& $v 'agg';if(!$agg){$agg='count'};$s=(& $v 'view')+' / "'+(& $v 'name')+'" = ';if($agg -ieq 'percent'){$s+=(& $v 'of')+' / '+(& $v 'per')}elseif($agg -ieq 'count'){$s+='patient count'}else{$s+=$agg+' of '+(& $v 'field')};if(& $v 'where'){$s+=' where '+(& $v 'where')};return $s}
 }
 return ('('+$t+' rule)')
}
function Find-CmbFileCandidates($Source,[object[]]$Files){
 # Loaded files (and sheets) that have every match column, newest first; fileName and sheet narrow them when given.
 $want=@($Source.match|ForEach-Object{ConvertTo-CmbHeaderKey $_});$out=@()
 foreach($f in $Files){
  $nm=[string](Get-P $f 'originalName' '');if($Source.fileName -and $nm -notlike $Source.fileName){continue}
  foreach($s in @(Get-P $f 'sheets' @())){
   if($Source.sheet -and [string]$s.name -ne $Source.sheet){continue}
   $have=@{};foreach($h in @(Get-P $s 'headers' @())){$have[(ConvertTo-CmbHeaderKey ([string]$h))]=$true}
   $ok=$want.Count -gt 0;foreach($w in $want){if(!$have.ContainsKey($w)){$ok=$false;break}}
   if($ok){$out+=@{token=[string]$f.token;file=$nm;sheet=[string]$s.name;headerRow=[int]$s.headerRow;rows=[int]$s.rows;loadedUtc=(ConvertTo-IsoText (Get-P $f 'loadedUtc' ''));headers=@(Get-P $s 'headers' @());fileDateUtc=(ConvertTo-IsoText (Get-P $f 'fileDateUtc' ''))}}
  }
 }
 return @($out)
}
function Get-CmbAssignments($Plan,[object[]]$Files,[hashtable]$Chosen=@{}){
 # The file for each source: the one chosen on the page if it still matches, otherwise the newest matching file not already used.
 $used=@{};$out=[ordered]@{}
 foreach($s in $Plan.sources){
  $cands=@(Find-CmbFileCandidates $s $Files);$pick=$null
  $want=[string]$Chosen[$s.key];if($want){foreach($c in $cands){if(($c.token+'|'+$c.sheet) -eq $want){$pick=$c}}}
  if(!$pick){foreach($c in $cands){if(!$used.ContainsKey($c.token+'|'+$c.sheet)){$pick=$c;break}};if(!$pick -and $cands.Count){$pick=$cands[0]}}
  if($pick){$used[$pick.token+'|'+$pick.sheet]=$true}
  $out[$s.key]=@{source=$s.id;candidates=$cands;pick=$pick}
 }
 return $out
}
function Get-CmbTablesFromAssignments($Assign){
 $t=@{};foreach($k in @($Assign.Keys)){$p=$Assign[$k].pick;if($p){$t[$k]=@{headers=@($p.headers);file=('"'+$p.file+'"'+$(if($p.sheet -and [IO.Path]::GetExtension($p.file) -ieq '.xlsx'){' sheet "'+$p.sheet+'"'}else{''}))}}};return $t
}
function Get-CmbCheck([object[]]$Rules,[object[]]$Files,[hashtable]$Chosen=@{}){
 # What the Rules and Output pages show: per-rule summaries and problems, plus which file each source would use.
 $plan0=Get-CmbPlan $Rules;$assign=Get-CmbAssignments $plan0 $Files $Chosen;$plan=Get-CmbPlan $Rules (Get-CmbTablesFromAssignments $assign)
 $rows=@(for($i=0;$i -lt $Rules.Count;$i++){
  $r=$Rules[$i];$type='';if($r -is [Collections.IDictionary]){$type=([string](Get-CmbValue $r 'rule' '')).Trim().ToLowerInvariant()}
  [ordered]@{index=($i+1);type=$type;summary=(Get-CmbRuleSummary $r);note=$(if($r -is [Collections.IDictionary]){Get-CmbText $r 'note'}else{''});json=$(if($r -is [Collections.IDictionary]){ConvertTo-CmbRuleText $r}else{''});problems=@($plan.problems|Where-Object{$_.rule -eq ($i+1)}|ForEach-Object{[ordered]@{severity=$_.severity;message=$_.message}})}
 })
 $srcs=@(foreach($s in $plan.sources){
  $a=$assign[$s.key];$pick=$a.pick
  [ordered]@{source=$s.id;label=$s.label;role=$s.role;required=$s.required;match=@($s.match);candidates=@($a.candidates|ForEach-Object{[ordered]@{value=($_.token+'|'+$_.sheet);file=$_.file;sheet=$_.sheet;rows=$_.rows;loadedUtc=$_.loadedUtc}});chosen=$(if($pick){$pick.token+'|'+$pick.sheet}else{''});file=$(if($pick){$pick.file}else{''});sheet=$(if($pick){$pick.sheet}else{''});rows=$(if($pick){$pick.rows}else{0})}
 })
 $shared=@{};foreach($s in $srcs){if($s.chosen){if($shared.ContainsKey($s.chosen)){$shared[$s.chosen]+=@($s.source)}else{$shared[$s.chosen]=@($s.source)}}}
 $general=@($plan.problems|Where-Object{$_.rule -eq 0}|ForEach-Object{[ordered]@{severity=$_.severity;message=$_.message}})
 foreach($k in @($shared.Keys)){if(@($shared[$k]).Count -gt 1){$general+=[ordered]@{severity='warning';message=('Sources '+(@($shared[$k]) -join ' and ')+' use the same file; make their match columns (or fileName) more specific.')}}}
 return [ordered]@{name=$plan.config.name;rules=$rows;general=$general;sources=$srcs;errors=$plan.errorCount;warnings=@($plan.problems|Where-Object{$_.severity -eq 'warning'}).Count;fields=@($plan.order);views=@($plan.views.Keys|ForEach-Object{$plan.views[$_].name})}
}
# --- Combine engine: reads the sources, matches rows to patients, combines and derives fields, and builds the views and audit tables.
function Get-CmbExprNames([string]$Text){
 # Names an expression refers to (not function names), for deciding which columns to read.
 try{$tk=Get-CmbTokens $Text}catch{return @()}
 $out=@();for($i=0;$i -lt $tk.Count;$i++){$t=$tk[$i];if($t.t -ne 'id'){continue};if(!$t.br -and $i+1 -lt $tk.Count -and $tk[$i+1].t -eq 'lp'){continue};$out+=[string]$t.v};return @($out)
}
function Invoke-CmbClean([string[]]$Values,[object[]]$Steps){
 # Cleaning steps applied in order to every value (trimmed first); '' is blank.
 $n=$Values.Length;$v=[string[]]::new($n);for($i=0;$i -lt $n;$i++){$x=$Values[$i];if($null -eq $x){$v[$i]=''}else{$v[$i]=$x.Trim()}}
 $ti=[Globalization.CultureInfo]::InvariantCulture.TextInfo
 foreach($s in $Steps){
  switch($s.op){
   'trim'{}
   'upper'{for($i=0;$i -lt $n;$i++){$v[$i]=$v[$i].ToUpperInvariant()}}
   'lower'{for($i=0;$i -lt $n;$i++){$v[$i]=$v[$i].ToLowerInvariant()}}
   'title'{for($i=0;$i -lt $n;$i++){$v[$i]=$ti.ToTitleCase($v[$i].ToLowerInvariant())}}
   'collapsespaces'{$rx=[regex]'\s+';for($i=0;$i -lt $n;$i++){$v[$i]=$rx.Replace($v[$i],' ')}}
   'digits'{$rx=[regex]'[^0-9]';for($i=0;$i -lt $n;$i++){$v[$i]=$rx.Replace($v[$i],'')}}
   'alnum'{$rx=[regex]'[^A-Za-z0-9]';for($i=0;$i -lt $n;$i++){$v[$i]=$rx.Replace($v[$i],'')}}
   'stripleadingzeros'{for($i=0;$i -lt $n;$i++){$x=$v[$i];if($x.StartsWith('0')){$y=$x.TrimStart('0');if($y.Length -eq 0){$y='0'};$v[$i]=$y}}}
   'pad'{$w=[int]$s.n;for($i=0;$i -lt $n;$i++){if($v[$i].Length -gt 0){$v[$i]=$v[$i].PadLeft($w,'0')}}}
   'left'{$w=[int]$s.n;for($i=0;$i -lt $n;$i++){if($v[$i].Length -gt $w){$v[$i]=$v[$i].Substring(0,$w)}}}
   'filldown'{$last='';for($i=0;$i -lt $n;$i++){if($v[$i].Length -eq 0){$v[$i]=$last}else{$last=$v[$i]}}}
   'removebracketid'{$rx=[regex]'\s*\[[^\]]*\]\s*$';for($i=0;$i -lt $n;$i++){$v[$i]=$rx.Replace($v[$i],'')}}
   'namekey'{$memo=@{};for($i=0;$i -lt $n;$i++){$x=$v[$i];if($x.Length -eq 0){continue};$y=$memo[$x];if($null -eq $y){$y=[CmbRt]::NameKey($x);$memo[$x]=$y};$v[$i]=$y}}
   'blankif'{$b=[string]$s.text;for($i=0;$i -lt $n;$i++){if([string]::Equals($v[$i],$b,[StringComparison]::OrdinalIgnoreCase)){$v[$i]=''}}}
   'replace'{$rx=[CmbRt]::Re([string]$s.rx);$w=[string]$s.with;for($i=0;$i -lt $n;$i++){if($v[$i].Length -gt 0){$v[$i]=$rx.Replace($v[$i],$w)}}}
   'extract'{$rx=[CmbRt]::Re([string]$s.rx);for($i=0;$i -lt $n;$i++){if($v[$i].Length -gt 0){$m=$rx.Match($v[$i]);if(!$m.Success){$v[$i]=''}elseif($m.Groups.Count -gt 1){$v[$i]=$m.Groups[1].Value}else{$v[$i]=$m.Value}}}}
   'map'{$h=$s.map;$else=[string]$s.else;for($i=0;$i -lt $n;$i++){$x=$v[$i];if($x.Length -eq 0){continue};$k=[CmbRt]::Key($x);if($h.ContainsKey($k)){$y=$h[$k];if($null -eq $y){$v[$i]=''}else{$v[$i]=[string]$y}}elseif($else -ieq 'blank'){$v[$i]=''}elseif($else -ine 'keep'){$v[$i]=$else}}}
  }
  if($s.op -notin @('trim','pad','left','blankif','map')){for($i=0;$i -lt $n;$i++){$v[$i]=$v[$i].Trim()}}
 }
 return ,$v
}
function ConvertTo-CmbTypedValues([object[]]$Values,[string]$Type,$Stat){
 # Typed values ($null = blank); a non-blank value the type cannot read becomes blank and is counted, with a few examples.
 # Text is only trimmed; other types remember each distinct text's result, since exports repeat values (dates, codes, flags).
 $n=$Values.Length;$out=[object[]]::new($n)
 if($Type -eq 'text'){for($i=0;$i -lt $n;$i++){$x=$Values[$i];if($null -eq $x){continue};if($x -is [string]){$t=$x.Trim();if($t.Length){$out[$i]=$t}}else{$out[$i]=[CmbRt]::Convert($x,'text')}};return ,$out}
 $memo=[Collections.Generic.Dictionary[string,object]]::new();$bad=[Collections.Generic.HashSet[string]]::new()
 for($i=0;$i -lt $n;$i++){
  $x=$Values[$i];if($null -eq $x){continue}
  if($x -is [string]){
   if($x.Length -eq 0){continue};$y=$null
   if($memo.TryGetValue($x,[ref]$y)){$out[$i]=$y;continue}
   if($bad.Contains($x)){$Stat.unreadable++;continue}
   $y=[CmbRt]::Convert($x,$Type)
   if($null -eq $y){if([string]::IsNullOrWhiteSpace($x)){continue};[void]$bad.Add($x);$Stat.unreadable++;if($Stat.examples.Count -lt 5){$t=$x.Trim();if(!$Stat.examples.Contains($t)){$Stat.examples.Add($t)}};continue}
   if($memo.Count -lt 200000){$memo[$x]=$y};$out[$i]=$y;continue
  }
  $y=[CmbRt]::Convert($x,$Type)
  if($null -eq $y){$Stat.unreadable++;if($Stat.examples.Count -lt 5){$t=[CmbRt]::Text($x);if(!$Stat.examples.Contains($t)){$Stat.examples.Add($t)}}}else{$out[$i]=$y}
 }
 return ,$out
}
function Get-CmbNumberText([double]$N){return $N.ToString('#,0',[Globalization.CultureInfo]::InvariantCulture)}
function Read-CmbSource($Plan,$Src,$Pick,$File,[object[]]$TableList,[hashtable]$TableMap){
 # One source: reads only the columns its rules use, applies the map rules in order, then its filters. Returns its fields and kept rows.
 $want=@{}
 foreach($m in $Src.maps){if($m.mode -eq 'column'){$want[(ConvertTo-CmbHeaderKey $m.column)]=$true};if($m.mode -eq 'columns'){foreach($c in $m.columns){$want[(ConvertTo-CmbHeaderKey $c)]=$true}};if($m.mode -eq 'expr'){foreach($x in (Get-CmbExprNames $m.expr)){$want[(ConvertTo-CmbHeaderKey $x)]=$true}}}
 foreach($f in $Src.filters){foreach($x in (Get-CmbExprNames $f.expr)){$want[(ConvertTo-CmbHeaderKey $x)]=$true}}
 $hr=$Src.headerRow;if($hr -le 0){$hr=[int]$Pick.headerRow}
 $tables=@(Read-CmbSheets ([string]$File.path) ([string]$Pick.sheet) $hr ([string[]]@($want.Keys)) ([string]$Pick.sheet) -CellText:($Src.read -eq 'text'))
 if($tables.Count -eq 0){throw ('Source "'+$Src.id+'": sheet "'+$Pick.sheet+'" was not found in '+[string]$File.originalName+'.')}
 $t=$tables[0];$n=[int]$t.rows
 $names=@{};$cols=[Collections.Generic.List[object]]::new()
 foreach($h in @($t.columns.Keys)){$a=$t.columns[$h];$o=New-Object object[] $n;[Array]::Copy([object[]]$a,$o,[Math]::Min($n,$a.Length));$cols.Add($o);$names[(ConvertTo-CmbHeaderKey $h)]=$cols.Count-1}
 $hdr=@{};foreach($h in @($t.headers)){$hdr[(ConvertTo-CmbHeaderKey $h)]=$h}
 $fields=@{};$effects=@{};$label=('"'+[string]$File.originalName+'"'+$(if([IO.Path]::GetExtension([string]$File.originalName) -ieq '.xlsx'){' sheet "'+$t.name+'"'}else{''}))
 $fileDate=$null;$fd=[string](Get-P $File 'fileDateUtc' '');if($fd){$d=ConvertTo-DateValue $fd;if($d){$fileDate=$d.ToLocalTime().Date}}
 foreach($m in $Src.maps){
  $stat=@{unreadable=0;examples=[Collections.Generic.List[string]]::new()};$raw=$null;$vals=$null;$clearedByClean=0
  switch($m.mode){
   'column'{$k=ConvertTo-CmbHeaderKey $m.column;if(!$names.ContainsKey($k) -or !$hdr.ContainsKey($k)){throw ('Source "'+$Src.id+'": column "'+$m.column+'" is not in '+$label+'.')};$raw=[string[]]$t.columns[$hdr[$k]]}
   'columns'{
    $parts=@(foreach($c in $m.columns){$k=ConvertTo-CmbHeaderKey $c;if(!$hdr.ContainsKey($k)){throw ('Source "'+$Src.id+'": column "'+$c+'" is not in '+$label+'.')};,([string[]]$t.columns[$hdr[$k]])})
    $raw=New-Object string[] $n;for($i=0;$i -lt $n;$i++){$sb=New-Object Text.StringBuilder;foreach($pa in $parts){$x=$pa[$i].Trim();if($x){if($sb.Length){[void]$sb.Append($m.sep)};[void]$sb.Append($x)}};$raw[$i]=$sb.ToString()}
   }
   'value'{$c=[CmbRt]::Text($m.value);$raw=New-Object string[] $n;for($i=0;$i -lt $n;$i++){$raw[$i]=$c}}
   'special'{
    $vals=New-Object object[] $n
    for($i=0;$i -lt $n;$i++){switch($m.special){'fileName'{$vals[$i]=[string]$File.originalName}'fileDate'{$vals[$i]=$fileDate}'rowNumber'{$vals[$i]=[double]($t.headerRow+1+$i)}'sourceLabel'{$vals[$i]=$Src.label}}}
   }
   'expr'{
    $c=New-CmbCompiled $m.expr (New-CmbExprContext $names $TableMap $Plan.config.asOf ('not a column of '+$label+' or a field mapped above'))
    $r=Invoke-CmbCompiled $c $cols.ToArray() $TableList $n;$vals=$r.values;if($r.errors.Count){$stat.unreadable+=$r.errors.Count}
    if($m.clean.Count){$raw=New-Object string[] $n;for($i=0;$i -lt $n;$i++){$raw[$i]=[CmbRt]::Text($vals[$i])};$vals=$null}
   }
  }
  if($null -ne $raw){
   $before=0;foreach($x in $raw){if($x.Length -gt 0 -and $x.Trim().Length -gt 0){$before++}}
   if($m.clean.Count){$raw=Invoke-CmbClean $raw $m.clean;$after=0;foreach($x in $raw){if($x.Length -gt 0){$after++}};$clearedByClean=$before-$after}
   $vals=ConvertTo-CmbTypedValues ([object[]]$raw) $m.type $stat
  }elseif($m.mode -eq 'expr'){$vals=ConvertTo-CmbTypedValues $vals $m.type $stat}
  $cols.Add($vals);$names[$m.key]=$cols.Count-1;$fields[$m.key]=$vals
  $filled=0;foreach($x in $vals){if($null -ne $x){$filled++}}
  $e='filled '+(Get-CmbNumberText $filled)+' of '+(Get-CmbNumberText $n)+' rows'
  if($clearedByClean -gt 0){$e+='; cleaning blanked '+(Get-CmbNumberText $clearedByClean)}
  if($stat.unreadable -gt 0){$e+='; '+(Get-CmbNumberText $stat.unreadable)+' value'+$(if($stat.unreadable -eq 1){''}else{'s'})+' could not be read as '+$m.type+$(if($stat.examples.Count){' (e.g. '+(@($stat.examples|ForEach-Object{'"'+$_+'"'}) -join ', ')+')'}else{''})+' and were left blank'}
  $effects[$m.ruleIndex]=$e
 }
 # Rows with every mapped field blank (spacer or footer rows) are not used.
 $keep=[bool[]]::new($n);$empty=0;$farr=[object[]]@(foreach($k in @($fields.Keys)){,$fields[$k]});$nf=$farr.Count
 for($i=0;$i -lt $n;$i++){$any=$false;for($j=0;$j -lt $nf;$j++){if($null -ne $farr[$j][$i]){$any=$true;break}};$keep[$i]=$any;if(!$any){$empty++}}
 $dropped=0
 foreach($f in $Src.filters){
  $c=New-CmbCompiled $f.expr (New-CmbExprContext $names $TableMap $Plan.config.asOf ('not a column of '+$label+' or a field of source "'+$Src.id+'"')) -AsBool
  $r=Invoke-CmbCompiled $c $cols.ToArray() $TableList $n;$d=0;$before=0
  for($i=0;$i -lt $n;$i++){if(!$keep[$i]){continue};$before++;$hit=[bool]$r.values[$i];if($f.keep -ne $hit){$keep[$i]=$false;$d++}}
  $dropped+=$d;$effects[$f.ruleIndex]=('removed '+(Get-CmbNumberText $d)+' of '+(Get-CmbNumberText $before)+' rows'+$(if($r.errors.Count){'; the expression failed on '+(Get-CmbNumberText $r.errors.Count)+' rows (treated as false)'}else{''}))
 }
 $notUnique=0
 if($Src.unique.Count){
  # A combination of the unique fields seen more than once in this source leaves out every row that has it.
  $ua=[object[]]@(foreach($u in $Src.unique){,$fields[$u.ToLowerInvariant()]});$cnt=[Collections.Generic.Dictionary[string,int]]::new();$kv=[string[]]::new($n)
  for($i=0;$i -lt $n;$i++){if(!$keep[$i]){continue};$parts=[string[]]::new($ua.Count);$ok=$true;for($j=0;$j -lt $ua.Count;$j++){$uk=[CmbRt]::Key($ua[$j][$i]);if($uk.Length -eq 0){$ok=$false;break};$parts[$j]=$uk};if(!$ok){continue};$k=[string]::Join('|',$parts);$kv[$i]=$k;if($cnt.ContainsKey($k)){$cnt[$k]++}else{$cnt[$k]=1}}
  for($i=0;$i -lt $n;$i++){if($keep[$i] -and $null -ne $kv[$i] -and $cnt[$kv[$i]] -gt 1){$keep[$i]=$false;$notUnique++}}
 }
 $rows=[Collections.Generic.List[int]]::new();for($i=0;$i -lt $n;$i++){if($keep[$i]){$rows.Add($i)}}
 return @{src=$Src;notUnique=$notUnique;n=$n;fields=$fields;rows=$rows;effects=$effects;file=$label;sheet=$t.name;headerRow=$t.headerRow;empty=$empty;dropped=$dropped;fileName=[string]$File.originalName;sha256=[string](Get-P $File 'sha256' '');token=[string]$File.token}
}
function Get-CmbRunTables($Plan,[hashtable]$Data){
 # lookup() tables in the order Get-CmbLookupTablesForCheck numbers them: inline lookup rules, then lookup sources (key -> row fields).
 $list=[Collections.Generic.List[object]]::new();$map=@{};$effects=@{}
 foreach($r in $Plan.rules){if($r -is [Collections.IDictionary] -and ([string](Get-CmbValue $r 'rule' '')).Trim() -ieq 'lookup'){$name=Get-CmbText $r 'table';if(!$name){continue};$tk=$name.ToLowerInvariant();$tb=@{};if($Plan.lookups.Contains($tk)){$tb=$Plan.lookups[$tk].table};$map[$tk]=@{index=$list.Count;fields=@()};$list.Add($tb)}}
 foreach($s in $Plan.sources){
  if($s.role -ne 'lookup'){continue};$tb=@{};$dup=0
  if($Data.ContainsKey($s.key)){
   $d=$Data[$s.key];$kk=$s.lookupKey.ToLowerInvariant();$ka=$d.fields[$kk]
   foreach($i in $d.rows){$k=[CmbRt]::Key($ka[$i]);if($k.Length -eq 0){continue};if($tb.ContainsKey($k)){$dup++;continue};$row=@{};foreach($f in @($s.fields.Keys)){$row[$s.fields[$f].name]=$d.fields[$f][$i]};$tb[$k]=$row}
   $effects[$s.ruleIndex]=('lookup table of '+(Get-CmbNumberText $tb.Count)+' keys'+$(if($dup){'; '+(Get-CmbNumberText $dup)+' rows repeat a key and were ignored'}else{''}))
  }
  $map[$s.key]=@{index=$list.Count;fields=@(foreach($f in @($s.fields.Keys)){$s.fields[$f].name})};$list.Add($tb)
 }
 if(!$Plan.ContainsKey('profileTables')){$Plan['profileTables']=Get-CmbProfileTables}
 foreach($k in @(Get-CmbProfileTableNames)){if(!$map.ContainsKey($k)){$map[$k]=@{index=$list.Count;fields=@('npi','provider','location','riskPool')}};$list.Add($Plan.profileTables[$k])}
 return @{list=$list.ToArray();map=$map;effects=$effects}
}
function Invoke-CmbEngine($Plan,[hashtable]$Picks,[scriptblock]$Progress={param($p,$s)}){
 # Picks: source key -> @{pick (candidate: token, sheet, headerRow);file (file record)}. Returns everything the workbook and the page need.
 $sw=[Diagnostics.Stopwatch]::StartNew();$effects=@{};$notes=[Collections.Generic.List[string]]::new()
 $data=@{};$tablesEarly=Get-CmbRunTables $Plan @{}
 $srcList=@($Plan.sources);$k=0
 # Lookup sources first, so patient sources' map expressions can use lookup().
 $readOrder=@($srcList|Where-Object{$_.role -eq 'lookup'})+@($srcList|Where-Object{$_.role -ne 'lookup'})
 foreach($s in $readOrder){
  $k++;$pk=$Picks[$s.key]
  if(!$pk -or !$pk.pick){if($s.required){throw ('No file is chosen for source "'+$s.id+'".')};$effects[$s.ruleIndex]='skipped: no file was chosen (the source is optional)';continue}
  & $Progress ([int](5+40*($k-1)/[Math]::Max(1,$readOrder.Count))) ('Reading '+$s.id+': '+[string]$pk.file.originalName)
  $tb=Get-CmbRunTables $Plan $data
  $d=Read-CmbSource $Plan $s $pk.pick $pk.file $tb.list $tb.map;$data[$s.key]=$d
  foreach($e in @($d.effects.Keys)){$effects[$e]=$d.effects[$e]}
 }
 $tb=Get-CmbRunTables $Plan $data;foreach($e in @($tb.effects.Keys)){$effects[$e]=$tb.effects[$e]}
 # Records: every kept row of every patient source, in source order then row order.
 & $Progress 48 'Matching rows to patients'
 $ps=@($srcList|Where-Object{$_.role -eq 'patients' -and $data.ContainsKey($_.key)});$sIndex=@{};for($i=0;$i -lt $ps.Count;$i++){$sIndex[$ps[$i].key]=$i}
 $N=0;foreach($s in $ps){$N+=$data[$s.key].rows.Count}
 $recSrc=New-Object int[] $N;$recRow=New-Object int[] $N;$r=0
 for($si=0;$si -lt $ps.Count;$si++){foreach($i in $data[$ps[$si].key].rows){$recSrc[$r]=$si;$recRow[$r]=$i;$r++}}
 $parent=[int[]]::new($N);for($i=0;$i -lt $N;$i++){$parent[$i]=$i}
 # Per key level: the group's key value at its root (kv) and the record it came from (kr, for the display text).
 $L=$Plan.keys.Count;$kv=@(for($j=0;$j -lt $L;$j++){,([string[]]::new($N))});$kr=@(for($j=0;$j -lt $L;$j++){$x=[int[]]::new($N);for($i=0;$i -lt $N;$i++){$x[$i]=-1};,$x});$mask=[int[]]::new($N);$hasKey=[bool[]]::new($N);$ambiguous=[bool[]]::new($N)
 $keyArrs=@(for($j=0;$j -lt $L;$j++){$key=$Plan.keys[$j];,@(for($si=0;$si -lt $ps.Count;$si++){if($key.sources -contains $ps[$si].id){,@(foreach($f in $key.fields){,$data[$ps[$si].key].fields[$f.ToLowerInvariant()]})}else{,$null}})})
 $keyConflicts=[Collections.Generic.List[object]]::new();$keyEffects=@{}
 for($lv=0;$lv -lt $L;$lv++){
  $key=$Plan.keys[$lv];$seen=[Collections.Generic.Dictionary[string,int]]::new();$joined=0;$withKey=0;$blocked=0;$arrs=$keyArrs[$lv];$kvl=$kv[$lv];$krl=$kr[$lv];$bit=1 -shl [Math]::Min(30,$lv)
  if($key.match -eq 'prefix'){
   # Rows of the key's sources join the single row of the "against" sources whose value begins with theirs; several such rows make the
   # row ambiguous (left out); none leaves it on its own (rows of these sources with the same value still join each other).
   $ag=@(for($si=0;$si -lt $ps.Count;$si++){if($key.against -contains $ps[$si].id){,@(foreach($f in $key.fields){,$data[$ps[$si].key].fields[$f.ToLowerInvariant()]})}else{,$null}})
   $cand=[Collections.Generic.Dictionary[string,int]]::new();$candFirst=[Collections.Generic.Dictionary[string,int]]::new()
   for($r=0;$r -lt $N;$r++){$fa=$ag[$recSrc[$r]];if($null -eq $fa){continue};$val=Get-CmbRecordKey $fa $recRow[$r];if(!$val){continue};if($cand.ContainsKey($val)){$cand[$val]++}else{$cand[$val]=1;$candFirst[$val]=$r}}
   $sorted=[string[]]@($cand.Keys);[Array]::Sort($sorted,[StringComparer]::Ordinal);$alone=0;$amb=0
   for($r=0;$r -lt $N;$r++){
    $fa=$arrs[$recSrc[$r]];if($null -eq $fa){continue};$val=Get-CmbRecordKey $fa $recRow[$r];if(!$val){continue};$withKey++
    $ix=[Array]::BinarySearch($sorted,$val,[StringComparer]::Ordinal);if($ix -lt 0){$ix=-bnot $ix};$total=0;$target=-1
    while($ix -lt $sorted.Length -and $sorted[$ix].StartsWith($val,[StringComparison]::Ordinal)){$total+=$cand[$sorted[$ix]];$target=$candFirst[$sorted[$ix]];$ix++}
    if($total -gt 1){$amb++;$ambiguous[$r]=$true;continue}
    $hasKey[$r]=$true
    $a=$r;while($parent[$a] -ne $a){$parent[$a]=$parent[$parent[$a]];$a=$parent[$a]}
    if($null -eq $kvl[$a]){$kvl[$a]=$val;$krl[$a]=$r}
    if($total -eq 1){$b=$target}elseif($seen.ContainsKey($val)){$b=$seen[$val]}else{$seen[$val]=$r;$alone++;continue}
    while($parent[$b] -ne $b){$parent[$b]=$parent[$parent[$b]];$b=$parent[$b]}
    if($a -eq $b){continue}
    $conf=-1;for($j=0;$j -lt $lv;$j++){$x=$kv[$j][$a];$y=$kv[$j][$b];if($null -ne $x -and $null -ne $y -and $x -ne $y){$conf=$j;break}}
    if($conf -ge 0){$blocked++;continue}
    $root=[Math]::Min($a,$b);$child=[Math]::Max($a,$b);$parent[$child]=$root;$joined++
    for($j=0;$j -le $lv;$j++){if($null -eq $kv[$j][$root] -and $null -ne $kv[$j][$child]){$kv[$j][$root]=$kv[$j][$child];$kr[$j][$root]=$kr[$j][$child]}}
    $mask[$root]=$mask[$root] -bor $mask[$child] -bor $bit
   }
   $keyEffects[$key.ruleIndex]=((Get-CmbNumberText $withKey)+' rows had this key; '+(Get-CmbNumberText $joined)+' joined by prefix or exact value; '+(Get-CmbNumberText $alone)+' matched no row of '+($key.against -join ', ')+$(if($amb){'; '+(Get-CmbNumberText $amb)+' left out because their value begins more than one row of '+($key.against -join ', ')}else{''})+$(if($blocked){'; '+(Get-CmbNumberText $blocked)+' joins refused because an earlier key differs'}else{''}))
   continue
  }
  for($r=0;$r -lt $N;$r++){
   $fa=$arrs[$recSrc[$r]];if($null -eq $fa){continue}
   $row=$recRow[$r]
   if($fa.Count -eq 1){$x=$fa[0][$row];if($null -eq $x){continue};$val=[CmbRt]::Key($x);if($val.Length -eq 0){continue}}
   else{$parts=[string[]]::new($fa.Count);$ok=$true;for($q=0;$q -lt $fa.Count;$q++){$t=[CmbRt]::Key($fa[$q][$row]);if($t.Length -eq 0){$ok=$false;break};$parts[$q]=$t};if(!$ok){continue};$val=[string]::Join('|',$parts)}
   $withKey++;$hasKey[$r]=$true
   $a=$r;while($parent[$a] -ne $a){$parent[$a]=$parent[$parent[$a]];$a=$parent[$a]}
   if($null -eq $kvl[$a]){$kvl[$a]=$val;$krl[$a]=$r}
   if(!$seen.ContainsKey($val)){$seen[$val]=$r;continue}
   $b=$seen[$val];while($parent[$b] -ne $b){$parent[$b]=$parent[$parent[$b]];$b=$parent[$b]}
   if($a -eq $b){continue}
   $conf=-1;for($j=0;$j -lt $lv;$j++){$x=$kv[$j][$a];$y=$kv[$j][$b];if($null -ne $x -and $null -ne $y -and $x -ne $y){$conf=$j;break}}
   if($conf -ge 0){$blocked++;if($keyConflicts.Count -lt 20000){$ra=$kr[$conf][$a];$rb=$kr[$conf][$b];$keyConflicts.Add([ordered]@{key=$key.label;value=(Get-CmbKeyDisplay $fa $row);earlierKey=$Plan.keys[$conf].label;first=(Get-CmbKeyDisplay $keyArrs[$conf][$recSrc[$rb]] $recRow[$rb]);second=(Get-CmbKeyDisplay $keyArrs[$conf][$recSrc[$ra]] $recRow[$ra]);sources=($ps[$recSrc[$b]].id+', '+$ps[$recSrc[$r]].id)})};continue}
   $root=[Math]::Min($a,$b);$child=[Math]::Max($a,$b);$parent[$child]=$root;$joined++
   for($j=0;$j -le $lv;$j++){if($null -eq $kv[$j][$root] -and $null -ne $kv[$j][$child]){$kv[$j][$root]=$kv[$j][$child];$kr[$j][$root]=$kr[$j][$child]}}
   $mask[$root]=$mask[$root] -bor $mask[$child] -bor $bit
  }
  $keyEffects[$key.ruleIndex]=((Get-CmbNumberText $withKey)+' rows had this key; '+(Get-CmbNumberText $joined)+' joins'+$(if($blocked){'; '+(Get-CmbNumberText $blocked)+' joins refused because an earlier key differs (see Key conflicts)'}else{''}))
 }
 foreach($e in @($keyEffects.Keys)){$effects[$e]=$keyEffects[$e]}
 # Patients in order of their first row; each patient's rows in record order.
 $pOf=[int[]]::new($N);$P=0;$rootPid=[int[]]::new($N);for($i=0;$i -lt $N;$i++){$rootPid[$i]=-1};$noKey=[int[]]::new($ps.Count)
 for($r=0;$r -lt $N;$r++){
  if(!$hasKey[$r]){$pOf[$r]=-1;$noKey[$recSrc[$r]]++;continue}
  $a=$r;while($parent[$a] -ne $a){$a=$parent[$a]};if($rootPid[$a] -ge 0){$pOf[$r]=$rootPid[$a]}else{$rootPid[$a]=$P;$pOf[$r]=$P;$P++}
 }
 $pStart=[int[]]::new($P+1);for($r=0;$r -lt $N;$r++){if($pOf[$r] -ge 0){$pStart[$pOf[$r]+1]++}};for($q=0;$q -lt $P;$q++){$pStart[$q+1]+=$pStart[$q]}
 $fill=[int[]]::new($P);$pRecs=[int[]]::new([Math]::Max(1,$pStart[$P]));for($r=0;$r -lt $N;$r++){$q=$pOf[$r];if($q -ge 0){$pRecs[$pStart[$q]+$fill[$q]]=$r;$fill[$q]++}}
 $pRoot=[int[]]::new($P);for($i=0;$i -lt $N;$i++){if($rootPid[$i] -ge 0){$pRoot[$rootPid[$i]]=$i}}
 # Built-in fields.
 $FV=@{};$LIN=@{}
 $keyText=[object[]]::new($P);$srcText=[object[]]::new($P);$rowsN=[object[]]::new($P);$matched=[object[]]::new($P);$mbText=@{};$mark=[int[]]::new($ps.Count);$sid=[string[]]@($ps|ForEach-Object{[string]$_.id})
 for($q=0;$q -lt $P;$q++){
  $a=$pRoot[$q];for($j=0;$j -lt $L;$j++){$rk=$kr[$j][$a];if($rk -ge 0){$keyText[$q]=(Get-CmbKeyDisplay $keyArrs[$j][$recSrc[$rk]] $recRow[$rk]);break}}
  $j0=$pStart[$q];$j1=$pStart[$q+1];$rowsN[$q]=[double]($j1-$j0)
  if($j1-$j0 -eq 1){$srcText[$q]=$sid[$recSrc[$pRecs[$j0]]]}
  else{$stamp=$q+1;$lp=[Collections.Generic.List[string]]::new();for($j=$j0;$j -lt $j1;$j++){$si=$recSrc[$pRecs[$j]];if($mark[$si] -ne $stamp){$mark[$si]=$stamp}};for($si=0;$si -lt $sid.Count;$si++){if($mark[$si] -eq $stamp){$lp.Add($sid[$si])}};$srcText[$q]=[string]::Join('; ',$lp)}
  $m=$mask[$a];if($m){if(!$mbText.ContainsKey($m)){$mbText[$m]=(@(for($j=0;$j -lt [Math]::Min(31,$L);$j++){if($m -band (1 -shl $j)){$Plan.keys[$j].label}}) -join '; ')};$matched[$q]=$mbText[$m]}
 }
 $FV['_key']=$keyText;$FV['_sources']=$srcText;$FV['_rows']=$rowsN;$FV['_matchedby']=$matched
 # Combined fields.
 $conflicts=[Collections.Generic.List[object]]::new();$conflictCount=0;$fk=0;$fieldNotes=[ordered]@{};$fieldKeys=@($Plan.fields.Keys)
 foreach($key in $fieldKeys){
  $fk++;if($fk % 5 -eq 1){& $Progress ([int](55+20*$fk/[Math]::Max(1,$fieldKeys.Count))) ('Combining fields ('+$fk+' of '+$fieldKeys.Count+')')}
  $f=$Plan.fields[$key];$res=Invoke-CmbCombineField $f $ps $data $P $pStart $pRecs $recSrc $recRow $keyText $conflicts
  $FV[$key]=$res.values;$LIN[$key]=$res.lineage;$conflictCount+=$res.conflicts
  $e=('filled for '+(Get-CmbNumberText $res.filled)+' of '+(Get-CmbNumberText $P)+' patients'+$(if($res.conflicts){'; sources disagreed for '+(Get-CmbNumberText $res.conflicts)+' patients (see Conflicts)'}else{''})+$(if($res.within){'; '+(Get-CmbNumberText $res.within)+' patients have different values in several rows of the chosen source (the first row was used)'}else{''}))
  if($f.ruleIndex -ge 0){$effects[$f.ruleIndex]=$e}else{$fieldNotes[$key]=$e}
 }
 # Derived fields, in order.
 & $Progress 76 'Computing derived fields'
 $colList=[Collections.Generic.List[object]]::new();$names=@{}
 foreach($key in @($FV.Keys)){$colList.Add($FV[$key]);$names[$key]=$colList.Count-1}
 foreach($d in $Plan.derives){
  $c=New-CmbCompiled $d.expr (New-CmbExprContext $names $tb.map $Plan.config.asOf 'not a field defined above')
  $r=Invoke-CmbCompiled $c $colList.ToArray() $tb.list $P;$vals=$r.values;$unread=0
  if($d.type){$stat=@{unreadable=0;examples=[Collections.Generic.List[string]]::new()};$vals=ConvertTo-CmbTypedValues $vals $d.type $stat;$unread=$stat.unreadable}
  else{for($q=0;$q -lt $P;$q++){$x=$vals[$q];if($x -is [string] -and $x.Length -eq 0){$vals[$q]=$null}}}
  $FV[$d.key]=$vals;$colList.Add($vals);$names[$d.key]=$colList.Count-1
  $filled=0;foreach($x in $vals){if($null -ne $x){$filled++}}
  $e='filled for '+(Get-CmbNumberText $filled)+' of '+(Get-CmbNumberText $P)+' patients'
  if($r.errors.Count){$ex=@($r.errors|Select-Object -First 3|ForEach-Object{[CmbRt]::Text($keyText[$_])});$e+='; the expression failed for '+(Get-CmbNumberText $r.errors.Count)+' patients (left blank; e.g. '+($ex -join ', ')+')'}
  if($unread){$e+='; '+(Get-CmbNumberText $unread)+' results could not be read as '+$d.type}
  $effects[$d.ruleIndex]=$e
 }
 # Patient filters.
 & $Progress 80 'Applying patient filters'
 $keepP=New-Object bool[] $P;for($q=0;$q -lt $P;$q++){$keepP[$q]=$true};$excluded=[Collections.Generic.List[object]]::new()
 foreach($f in $Plan.patientFilters){
  $c=New-CmbCompiled $f.expr (New-CmbExprContext $names $tb.map $Plan.config.asOf 'not a field') -AsBool
  $r=Invoke-CmbCompiled $c $colList.ToArray() $tb.list $P;$dn=0;$before=0
  for($q=0;$q -lt $P;$q++){if(!$keepP[$q]){continue};$before++;if($f.keep -ne [bool]$r.values[$q]){$keepP[$q]=$false;$dn++;if($excluded.Count -lt 200000){$excluded.Add(@($q,$f.ruleIndex))}}}
  $effects[$f.ruleIndex]=('removed '+(Get-CmbNumberText $dn)+' of '+(Get-CmbNumberText $before)+' patients'+$(if($r.errors.Count){'; the expression failed for '+(Get-CmbNumberText $r.errors.Count)+' (treated as false)'}else{''}))
 }
 $kept=[Collections.Generic.List[int]]::new();for($q=0;$q -lt $P;$q++){if($keepP[$q]){$kept.Add($q)}}
 # Sort.
 $order=$kept.ToArray()
 if($Plan.sorts.Count -and $order.Length -gt 1){
  & $Progress 82 'Sorting'
  $order=Get-CmbSortedOrder $order @(foreach($s in $Plan.sorts){,@($FV[$s.key],[bool]$s.desc)})
  foreach($s in $Plan.sorts){$effects[$s.ruleIndex]='applied'}
 }
 # Output columns.
 $outCols=@()
 if($Plan.outputs.Count){$outCols=@(foreach($o in $Plan.outputs){@{key=$o.key;header=$o.header;format=$o.format;width=$o.width;ruleIndex=$o.ruleIndex}})}
 else{$outCols=@(foreach($fname in $Plan.order){@{key=$fname.ToLowerInvariant();header=$fname;format='';width=$null;ruleIndex=-1}})+@(@{key='_sources';header='_sources';format='';width=$null;ruleIndex=-1})}
 foreach($o in $Plan.outputs){$vals=$FV[$o.key];$filled=0;foreach($q in $order){if($null -ne $vals[$q]){$filled++}};$effects[$o.ruleIndex]=('filled for '+(Get-CmbNumberText $filled)+' of '+(Get-CmbNumberText $order.Length)+' patients')}
 # Views.
 & $Progress 85 'Building views'
 $views=@(foreach($vk in @($Plan.views.Keys)){$v=$Plan.views[$vk];$t=Get-CmbViewTable $v $FV $names $colList $tb $Plan $order;$effects[$v.ruleIndex]=((Get-CmbNumberText $t.rows.Count)+' groups from '+(Get-CmbNumberText $t.patients)+' patients');$t})
 # Source statistics.
 $perSource=[int[]]::new($ps.Count);$lastP=[int[]]::new($ps.Count);for($i=0;$i -lt $ps.Count;$i++){$lastP[$i]=-1}
 for($q=0;$q -lt $P;$q++){if(!$keepP[$q]){continue};for($j=$pStart[$q];$j -lt $pStart[$q+1];$j++){$si=$recSrc[$pRecs[$j]];if($lastP[$si] -ne $q){$lastP[$si]=$q;$perSource[$si]++}}}
 $sources=@(foreach($s in $srcList){
  if(!$data.ContainsKey($s.key)){[ordered]@{source=$s.id;label=$s.label;role=$s.role;file='';sheet='';headerRow='';sha256='';rowsRead=0;emptyRows=0;filtered=0;noKey=0;rowsUsed=0;patients=0};continue}
  $d=$data[$s.key];$si=-1;if($sIndex.ContainsKey($s.key)){$si=$sIndex[$s.key]}
  $pc=0;if($si -ge 0){$pc=$perSource[$si]}
  $nk=$(if($si -ge 0){$noKey[$si]}else{0})
  if(!$effects.ContainsKey($s.ruleIndex)){$effects[$s.ruleIndex]=('read '+(Get-CmbNumberText $d.n)+' rows from '+$d.file+$(if($d.notUnique){'; '+(Get-CmbNumberText $d.notUnique)+' left out because their '+($s.unique -join ' + ')+' appears more than once'}else{''})+'; '+(Get-CmbNumberText ($d.rows.Count-$nk))+' used'+$(if($s.role -eq 'patients'){' for '+(Get-CmbNumberText $pc)+' patients'}else{''}))}
  [ordered]@{source=$s.id;label=$s.label;role=$s.role;file=$d.fileName;sheet=$d.sheet;headerRow=$d.headerRow;sha256=$d.sha256;rowsRead=$d.n;emptyRows=$d.empty;filtered=($d.dropped+$d.notUnique);noKey=$nk;rowsUsed=($d.rows.Count-$nk);patients=$pc;token=$d.token}
 })
 $noKeyTotal=0;foreach($x in $noKey){$noKeyTotal+=$x}
 return @{plan=$Plan;fieldNotes=$fieldNotes;F=$FV;LIN=$LIN;P=$P;order=$order;outCols=$outCols;views=$views;effects=$effects;sources=$sources;conflicts=$conflicts;conflictCount=$conflictCount;keyConflicts=$keyConflicts;excluded=$excluded;keyText=$keyText;records=$N;noKey=$noKeyTotal;seconds=[Math]::Round($sw.Elapsed.TotalSeconds,1)}
}
function Get-CmbRecordKey([object[]]$Arrays,[int]$Row){if($Arrays.Count -eq 1){return [CmbRt]::Key($Arrays[0][$Row])};$p=[string[]]::new($Arrays.Count);for($i=0;$i -lt $Arrays.Count;$i++){$t=[CmbRt]::Key($Arrays[$i][$Row]);if($t.Length -eq 0){return ''};$p[$i]=$t};return [string]::Join('|',$p)}
function Get-CmbKeyDisplay([object[]]$Arrays,[int]$Row){$p=[string[]]::new($Arrays.Count);for($i=0;$i -lt $Arrays.Count;$i++){$p[$i]=[CmbRt]::Text($Arrays[$i][$Row]).Trim()};return [string]::Join(' / ',$p)}
function Get-CmbSortedOrder([int[]]$Order,[object[]]$Keys){
 # Stable multi-key sort without objects: each patient gets one ordinal-sortable string. Numbers and dates become fixed-width text,
 # descending keys have their characters inverted, blanks sort last either way, and the original position breaks ties.
 $n=$Order.Length;$parts=@(foreach($k in $Keys){$vals=$k[0];$kind=Get-CmbSortKind @(foreach($q in $Order){$vals[$q]});,@($vals,$kind,[bool]$k[1])})
 $sk=[string[]]::new($n);$sb=[Text.StringBuilder]::new()
 for($i=0;$i -lt $n;$i++){
  $q=$Order[$i];[void]$sb.Clear()
  foreach($p in $parts){
   $v=$p[0][$q];if($null -eq $v){[void]$sb.Append('2');[void]$sb.Append([char]0);continue};[void]$sb.Append('1')
   if($p[1] -eq 'number'){$b=[BitConverter]::DoubleToInt64Bits([double]$v);if($b -lt 0){$b=-bnot $b}else{$b=$b -bor [long]::MinValue};$t=$b.ToString('X16')}
   elseif($p[1] -eq 'date'){$t=([DateTime]$v).Ticks.ToString('D19')}
   else{$t=[CmbRt]::Text($v).Trim().ToUpperInvariant()+[char]0}
   if($p[2]){$ca=$t.ToCharArray();for($j=0;$j -lt $ca.Length;$j++){$ca[$j]=[char](65535-[int]$ca[$j])};$t=[string]::new($ca)}
   [void]$sb.Append($t)
  }
  [void]$sb.Append($i.ToString('D10'));$sk[$i]=$sb.ToString()
 }
 $items=[int[]]$Order.Clone();[Array]::Sort($sk,$items,[StringComparer]::Ordinal);return ,$items
}
function Get-CmbSortKind([object[]]$Values){$k='';foreach($v in $Values){if($null -eq $v){continue};$t=$(if($v -is [DateTime]){'date'}elseif([CmbRt]::IsNum($v)){'number'}else{'text'});if(!$k){$k=$t}elseif($k -ne $t){return 'text'}};if(!$k){return 'text'};return $k}
function ConvertTo-CmbSortValue($V,[string]$Kind,[bool]$Desc){
 # Same-type keys so Sort-Object compares like with like; blanks sort last either way.
 if($Kind -eq 'date'){if($null -eq $V){if($Desc){return [DateTime]::MinValue};return [DateTime]::MaxValue};return [DateTime]$V}
 if($Kind -eq 'number'){if($null -eq $V){if($Desc){return [double]::NegativeInfinity};return [double]::PositiveInfinity};return [double]$V}
 if($null -eq $V){if($Desc){return ''};return [string][char]0xFFFF};return [CmbRt]::Text($V).ToUpperInvariant()
}
function Invoke-CmbCombineField($F,[object[]]$Ps,[hashtable]$Data,[int]$P,[int[]]$PStart,[int[]]$PRecs,[int[]]$RecSrc,[int[]]$RecRow,[object[]]$KeyText,$Conflicts){
 # One combined field for every patient: the value, which source(s) gave it, and conflicts where sources disagree with the chosen value.
 $SC=$Ps.Count;$rank=[int[]]::new($SC);$vals=[object[]]::new($SC);$by=[object[]]::new($SC);$ids=[string[]]::new($SC)
 for($s=0;$s -lt $SC;$s++){$ids[$s]=[string]$Ps[$s].id;$rank[$s]=[Array]::IndexOf([string[]]@($F.from),$ids[$s]);$fl=$Data[$Ps[$s].key].fields;if($fl.ContainsKey($F.ofKey)){$vals[$s]=$fl[$F.ofKey]};if($F.byKey -and $fl.ContainsKey($F.byKey)){$by[$s]=$fl[$F.byKey]}}
 $how=[string]$F.how;$type=[string]$F.type
 # 1 priority, 2 latest/earliest, 3 max/min, 4 mostCommon: one row's value is chosen. 5: an aggregate of every row.
 $mode=5;if($how -eq 'priority'){$mode=1}elseif($how -eq 'latest' -or $how -eq 'earliest'){$mode=2}elseif($how -eq 'max' -or $how -eq 'min'){$mode=3}elseif($how -eq 'mostCommon'){$mode=4}
 $agg=0;if($mode -eq 5){$agg=[Array]::IndexOf([string[]]@('sum','avg','count','countDistinct','list','any','all','exists'),$how)+1}
 $convert=($type -and $mode -le 4) -or ($type -and $agg -eq 5)
 $out=[object[]]::new($P);$lin=[string[]]::new($P);$filled=0;$nConf=0;$nWithin=0;$report=($F.conflicts -eq 'report' -and $mode -ne 3 -and $mode -ne 5);$sign=1;if($how -eq 'earliest' -or $how -eq 'min'){$sign=-1}
 $mark=[int[]]::new($SC)
 for($q=0;$q -lt $P;$q++){
  $j0=$PStart[$q];$j1=$PStart[$q+1]
  if($mode -le 4 -and $j1-$j0 -eq 1){
   # One row: its value, if its source is used.
   $r=$PRecs[$j0];$s=$RecSrc[$r];if($rank[$s] -lt 0){continue};$a=$vals[$s];if($null -eq $a){continue};$v=$a[$RecRow[$r]];if($null -eq $v){continue}
   if($convert){$v=[CmbRt]::Convert($v,$type);if($null -eq $v){continue}};$out[$q]=$v;$lin[$q]=$ids[$s];$filled++;continue
  }
  $best=$null;$bestS=-1;$bestRank=[int]::MaxValue;$bestD=$null
  if($mode -eq 1){
   for($j=$j0;$j -lt $j1;$j++){$r=$PRecs[$j];$s=$RecSrc[$r];$rk=$rank[$s];if($rk -lt 0 -or $rk -ge $bestRank){continue};$a=$vals[$s];if($null -eq $a){continue};$v=$a[$RecRow[$r]];if($null -eq $v){continue};$best=$v;$bestS=$s;$bestRank=$rk}
  }elseif($mode -eq 2){
   for($j=$j0;$j -lt $j1;$j++){
    $r=$PRecs[$j];$s=$RecSrc[$r];$rk=$rank[$s];if($rk -lt 0){continue};$a=$vals[$s];if($null -eq $a){continue};$v=$a[$RecRow[$r]];if($null -eq $v){continue}
    $d=$null;$ba=$by[$s];if($null -ne $ba){$d=$ba[$RecRow[$r]];if($null -ne $d -and !($d -is [DateTime])){$d=[CmbRt]::Date($d)}}
    $take=$false
    if($bestS -lt 0){$take=$true}
    elseif($null -ne $d){if($null -eq $bestD){$take=$true}else{$c=([DateTime]$d).CompareTo([DateTime]$bestD)*$sign;if($c -gt 0 -or ($c -eq 0 -and $rk -lt $bestRank)){$take=$true}}}
    elseif($null -eq $bestD -and $rk -lt $bestRank){$take=$true}
    if($take){$best=$v;$bestS=$s;$bestRank=$rk;$bestD=$d}
   }
  }elseif($mode -eq 3){
   for($j=$j0;$j -lt $j1;$j++){$r=$PRecs[$j];$s=$RecSrc[$r];if($rank[$s] -lt 0){continue};$a=$vals[$s];if($null -eq $a){continue};$v=$a[$RecRow[$r]];if($null -eq $v){continue};if($bestS -lt 0){$best=$v;$bestS=$s;continue};$c=[CmbRt]::Cmp($v,$best);if($null -ne $c -and [int]$c*$sign -gt 0){$best=$v;$bestS=$s}}
  }elseif($mode -eq 4){
   $cnt=@{};$first=@{};$orig=@{}
   for($j=$j0;$j -lt $j1;$j++){$r=$PRecs[$j];$s=$RecSrc[$r];$rk=$rank[$s];if($rk -lt 0){continue};$a=$vals[$s];if($null -eq $a){continue};$v=$a[$RecRow[$r]];if($null -eq $v){continue};$k=[CmbRt]::Key($v);if($cnt.ContainsKey($k)){$cnt[$k]++;if($rk -lt $first[$k][0]){$first[$k]=@($rk,$s)}}else{$cnt[$k]=1;$first[$k]=@($rk,$s);$orig[$k]=$v}}
   $bk=$null;foreach($k in @($cnt.Keys)){if($null -eq $bk -or $cnt[$k] -gt $cnt[$bk] -or ($cnt[$k] -eq $cnt[$bk] -and $first[$k][0] -lt $first[$bk][0])){$bk=$k}}
   if($null -ne $bk){$best=$orig[$bk];$bestS=$first[$bk][1]}
  }else{
   # 1 sum, 2 avg, 3 count, 4 countDistinct, 5 list, 6 any, 7 all, 8 exists
   $n=0;$sum=0.0;$set=$null;$list=$null;$tr=0;$used=0;$stamp=$q+1
   for($j=$j0;$j -lt $j1;$j++){
    $r=$PRecs[$j];$s=$RecSrc[$r];$rk=$rank[$s];if($rk -lt 0){continue}
    if($agg -ne 8){$a=$vals[$s];if($null -eq $a){continue};$v=$a[$RecRow[$r]];if($null -eq $v){continue}}
    if($mark[$s] -ne $stamp){$mark[$s]=$stamp;$used++}
    if($agg -le 2){$x=[CmbRt]::Num($v);if($null -ne $x){$sum+=[double]$x;$n++}}
    elseif($agg -eq 3 -or $agg -eq 8){$n++}
    elseif($agg -eq 4){if($null -eq $set){$set=@{}};$set[[CmbRt]::Key($v)]=$true}
    elseif($agg -eq 5){if($null -eq $set){$set=@{};$list=[Collections.Generic.List[object[]]]::new()};$k=[CmbRt]::Key($v);if(!$set.ContainsKey($k)){$set[$k]=$true;$list.Add([object[]]@($rk,$j,[CmbRt]::Text($v).Trim()))}}
    else{$n++;if([CmbRt]::Truthy($v)){$tr++}}
   }
   if($agg -eq 1){if($n){$best=$sum}}elseif($agg -eq 2){if($n){$best=$sum/$n}}elseif($agg -eq 3){$best=[double]$n}elseif($agg -eq 4){$best=[double]0;if($set){$best=[double]$set.Count}}
   elseif($agg -eq 5){if($list){if($list.Count -gt 1){$list.Sort([Comparison[object[]]]{param($x,$y)$c=([int]$x[0]).CompareTo([int]$y[0]);if($c -eq 0){$c=([int]$x[1]).CompareTo([int]$y[1])};$c})};$parts=[string[]]::new($list.Count);for($i=0;$i -lt $list.Count;$i++){$parts[$i]=[string]$list[$i][2]};$best=[string]::Join($F.sep,$parts)}}
   elseif($agg -eq 6){if($n){$best=($tr -gt 0)}}elseif($agg -eq 7){if($n){$best=($tr -eq $n)}}elseif($agg -eq 8){$best=($n -gt 0)}
   if($used){if($used -eq 1){for($s=0;$s -lt $SC;$s++){if($mark[$s] -eq $stamp){$lin[$q]=$ids[$s];break}}}else{$lp=[Collections.Generic.List[string]]::new();for($s=0;$s -lt $SC;$s++){if($mark[$s] -eq $stamp){$lp.Add($ids[$s])}};$lin[$q]=[string]::Join(',',$lp)}}
   if($convert -and $null -ne $best){$best=[CmbRt]::Convert($best,$type)}
   $out[$q]=$best;if($null -ne $best){$filled++};continue
  }
  if($bestS -lt 0){continue}
  if($convert){$best=[CmbRt]::Convert($best,$type);if($null -eq $best){continue}}
  $out[$q]=$best;$lin[$q]=$ids[$bestS];$filled++
  if($report){
   # Only another source disagreeing is a conflict; several values within the chosen source (repeated rows) are counted separately.
   $others=$null;$within=$false
   for($j=$j0;$j -lt $j1;$j++){
    $r=$PRecs[$j];$s=$RecSrc[$r];if($rank[$s] -lt 0){continue};$a=$vals[$s];if($null -eq $a){continue};$v=$a[$RecRow[$r]];if($null -eq $v){continue}
    if($v -is [string] -and $best -is [string]){$same=[string]::Equals($v.Trim(),$best.Trim(),[StringComparison]::OrdinalIgnoreCase)}else{$same=[CmbRt]::Eq($v,$best)}
    if($same){continue};if($s -eq $bestS){$within=$true;continue}
    if($null -eq $others){$others=[ordered]@{}};$others[([CmbRt]::Text($v)+' ['+$ids[$s]+']')]=$true
   }
   if($within){$nWithin++}
   if($null -ne $others){$nConf++;if($Conflicts.Count -lt 100000){$Conflicts.Add([ordered]@{patient=[CmbRt]::Text($KeyText[$q]);field=$F.name;chosen=[CmbRt]::Text($best);from=$ids[$bestS];others=(@($others.Keys) -join '; ');how=$how;ruleIndex=$F.ruleIndex})}}
  }
 }
 return @{values=$out;lineage=$lin;filled=$filled;conflicts=$nConf;within=$nWithin}
}
function Get-CmbViewTable($View,[hashtable]$FV,[hashtable]$Names,$ColList,$Tables,$Plan,[int[]]$Order){
 # One summary view: patients (in output order, after patient filters) grouped by the view's fields, one column per measure, plus a total.
 $P=$FV['_rows'].Length;$cols=$ColList.ToArray()
 $inView=[bool[]]::new($P);foreach($q in $Order){$inView[$q]=$true}
 if($View.where){$c=New-CmbCompiled $View.where (New-CmbExprContext $Names $Tables.map $Plan.config.asOf 'not a field') -AsBool;$r=Invoke-CmbCompiled $c $cols $Tables.list $P;for($q=0;$q -lt $P;$q++){if(!$r.values[$q]){$inView[$q]=$false}}}
 $ms=@($View.measures);$MC=$ms.Count
 $mWhere=[object[]]@(foreach($m in $ms){if($m.where){$c=New-CmbCompiled $m.where (New-CmbExprContext $Names $Tables.map $Plan.config.asOf 'not a field') -AsBool;,((Invoke-CmbCompiled $c $cols $Tables.list $P).values)}else{,$null}})
 $mVals=[object[]]@(foreach($m in $ms){if($m.field){,$FV[$m.field.ToLowerInvariant()]}else{,$null}})
 # 1 count, 2 sum, 3 avg, 4 min, 5 max, 6 countDistinct, 0 percent (computed after)
 $mAgg=[int[]]@(foreach($m in $ms){[Array]::IndexOf([string[]]@('percent','count','sum','avg','min','max','countDistinct'),[string]$m.agg)})
 $gk=@($View.groupBy|ForEach-Object{$_.ToLowerInvariant()});$garr=[object[]]@(foreach($g in $gk){,$FV[$g]});$GC=$garr.Count
 $index=[Collections.Generic.Dictionary[string,int]]::new();$labels=[Collections.Generic.List[object]]::new();$accs=[Collections.Generic.List[object]]::new()
 $total=New-CmbAccumulators $ms;$patients=0;$lab=[string[]]::new($GC)
 foreach($q in $Order){
  if(!$inView[$q]){continue};$patients++
  for($g=0;$g -lt $GC;$g++){$x=[CmbRt]::Text($garr[$g][$q]).Trim();if($x.Length -eq 0){$x='(blank)'};$lab[$g]=$x}
  $k=[string]::Join([string][char]31,$lab).ToUpperInvariant();$gi=0
  if(!$index.TryGetValue($k,[ref]$gi)){$gi=$labels.Count;$index[$k]=$gi;$labels.Add([string[]]$lab.Clone());$accs.Add((New-CmbAccumulators $ms))}
  $acc=$accs[$gi]
  for($i=0;$i -lt $MC;$i++){
   $w=$mWhere[$i];if($null -ne $w -and !$w[$q]){continue}
   $ag=$mAgg[$i];if($ag -eq 0){continue};$a1=$acc[$i];$a2=$total[$i]
   if($ag -eq 1){$a1.n++;$a2.n++;continue}
   $v=$mVals[$i][$q];if($null -eq $v){continue}
   if($ag -le 3){$x=[CmbRt]::Num($v);if($null -ne $x){$a1.sum+=[double]$x;$a1.n++;$a2.sum+=[double]$x;$a2.n++};continue}
   if($ag -eq 6){if($null -eq $a1.set){$a1.set=@{}};if($null -eq $a2.set){$a2.set=@{}};$kk=[CmbRt]::Key($v);$a1.set[$kk]=$true;$a2.set[$kk]=$true;continue}
   $sgn=$(if($ag -eq 4){-1}else{1})
   foreach($ax in @($a1,$a2)){if($null -eq $ax.best){$ax.best=$v}else{$c=[CmbRt]::Cmp($v,$ax.best);if($null -ne $c -and [int]$c*$sgn -gt 0){$ax.best=$v}}}
  }
 }
 $rows=@(for($gi=0;$gi -lt $labels.Count;$gi++){,(@($labels[$gi])+(Get-CmbMeasureValues $ms $accs[$gi]))})
 # Sort: by a measure or group field (blanks last), otherwise by the group labels.
 $headers=@($View.groupBy)+@($ms|ForEach-Object{$_.name});$si=-1;if($View.sortBy){for($i=0;$i -lt $headers.Count;$i++){if($headers[$i] -ieq $View.sortBy){$si=$i}}}
 if($rows.Count -gt 1){
  $ord=[int[]]@(0..($rows.Count-1));$keys=@()
  if($si -ge 0){$col=[object[]]@(foreach($r in $rows){$r[$si]});$keys+=,@($col,[bool]$View.desc)}
  for($g=0;$g -lt $GC;$g++){$col=[object[]]@(foreach($r in $rows){[string]$r[$g]});$keys+=,@($col,$(if($si -ge 0){$false}else{[bool]$View.desc}))}
  $ord=Get-CmbSortedOrder $ord $keys;$rows=@(foreach($i in $ord){,$rows[$i]})
 }
 $totalRow=$null;if($View.total){$tl=@(for($i=0;$i -lt $GC;$i++){if($i -eq 0){'Total'}else{''}});$totalRow=@($tl)+(Get-CmbMeasureValues $ms $total)}
 return @{name=$View.name;ruleIndex=$View.ruleIndex;headers=$headers;groupCount=$GC;formats=@($ms|ForEach-Object{$mf=$_.format;if(!$mf){$mf=$(if($_.agg -eq 'percent'){'0.0%'}elseif($_.agg -in @('count','countDistinct')){'#,##0'}else{''})};$mf});rows=$rows;total=$totalRow;patients=$patients}
}
function New-CmbAccumulators([object[]]$Measures){return ,@(foreach($m in $Measures){@{n=0;sum=0.0;best=$null;set=$null}})}
function Get-CmbMeasureValues([object[]]$Measures,[object[]]$Acc){
 $vals=@(for($i=0;$i -lt $Measures.Count;$i++){$m=$Measures[$i];$a=$Acc[$i]
  switch($m.agg){'count'{[double]$a.n}'sum'{if($a.n){$a.sum}else{$null}}'avg'{if($a.n){$a.sum/$a.n}else{$null}}'min'{$a.best}'max'{$a.best}'countDistinct'{if($a.set){[double]$a.set.Count}else{[double]0}}'percent'{$null}}})
 for($i=0;$i -lt $Measures.Count;$i++){$m=$Measures[$i];if($m.agg -ne 'percent'){continue};$o=-1;$p=-1;for($j=0;$j -lt $Measures.Count;$j++){if($Measures[$j].name -ieq $m.of){$o=$j};if($Measures[$j].name -ieq $m.per){$p=$j}};if($o -ge 0 -and $p -ge 0){$vals[$i]=[CmbRt]::Div($vals[$o],$vals[$p])}}
 return ,$vals
}
# --- Combine output: the workbook (patient sheet, views, audit sheets), run records and the worker.
function Get-CmbCellValue($V,[string]$Booleans){
 if($null -eq $V){return $null}
 if($V -is [bool]){switch($Booleans){'Yes/No'{if($V){return 'Yes'};return 'No'}'TRUE/FALSE'{return [bool]$V}'1/0'{if($V){return [double]1};return [double]0}default{if($V){return 'Y'};return 'N'}}}
 if($V -is [string]){if($V.Length -gt 32000){return $V.Substring(0,32000)};return $V}
 return $V
}
function Format-CmbDisplay($V,[string]$Booleans,[string]$Format=''){
 # Text for the page preview: dates as M/d/yyyy, percent formats as percents, other numbers with up to two decimals.
 if($null -eq $V){return ''};if($V -is [bool]){return [string](Get-CmbCellValue $V $Booleans)}
 if($V -is [DateTime]){return [CmbRt]::Text($V)}
 if([CmbRt]::IsNum($V)){$d=[double]$V;if($Format -like '*%*'){return ($d*100).ToString('0.0',[Globalization.CultureInfo]::InvariantCulture)+'%'};if([Math]::Abs($d-[Math]::Round($d)) -lt 1e-9){return ([Math]::Round($d)).ToString('#,0',[Globalization.CultureInfo]::InvariantCulture)};return $d.ToString('#,0.##',[Globalization.CultureInfo]::InvariantCulture)}
 return [string]$V
}
function Get-CmbSheetName([string]$Name,[hashtable]$Used){
 $n=($Name -replace '[\[\]:*?/\\]',' ').Trim();if(!$n){$n='Sheet'};if($n.Length -gt 31){$n=$n.Substring(0,31).Trim()}
 $b=$n;$k=2;while($Used.ContainsKey($n.ToLowerInvariant())){$s=' ('+$k+')';$n=$b.Substring(0,[Math]::Min($b.Length,31-$s.Length))+$s;$k++};$Used[$n.ToLowerInvariant()]=$true;return $n
}
function Add-CmbSheet($Pkg,[hashtable]$Used,[string]$Name,[string[]]$Headers,$Rows,[string[]]$Formats=$null,[object[]]$Widths=$null,[int]$BoldLastRow=0){
 # One sheet: a styled header row, the rows (one bulk load), number formats and widths per column, filter and frozen header.
 try{Add-Type -AssemblyName System.Drawing -ErrorAction Stop}catch{}
 $ws=$Pkg.Workbook.Worksheets.Add((Get-CmbSheetName $Name $Used));$nc=$Headers.Count;if($nc -eq 0){return $ws}
 $list=[Collections.Generic.List[object[]]]::new();$list.Add([object[]]$Headers)
 if($Rows -is [Collections.Generic.List[object[]]]){$list.AddRange($Rows)}else{foreach($r in $Rows){$list.Add([object[]]$r)}}
 $null=$ws.Cells[1,1].LoadFromArrays($list);$nr=$list.Count
 $h=$ws.Cells[1,1,1,$nc];$h.Style.Font.Bold=$true;$h.Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid;$h.Style.Fill.BackgroundColor.SetColor([Drawing.Color]::FromArgb(23,54,93));$h.Style.Font.Color.SetColor([Drawing.Color]::White);$h.Style.WrapText=$true
 for($c=0;$c -lt $nc;$c++){
  $fmt='';if($Formats -and $c -lt $Formats.Count){$fmt=[string]$Formats[$c]}
  if(!$fmt -and $nr -gt 1){$hasTime=$false;$isDate=$false;for($r=1;$r -lt [Math]::Min($nr,2001);$r++){$v=$list[$r][$c];if($v -is [DateTime]){$isDate=$true;if(([DateTime]$v).TimeOfDay.Ticks -ne 0){$hasTime=$true}}};if($isDate){$fmt=$(if($hasTime){'m/d/yy h:mm'}else{'mm-dd-yy'})}}
  if($fmt -and $nr -gt 1){$ws.Cells[2,($c+1),$nr,($c+1)].Style.Numberformat.Format=$fmt}
  $w=$null;if($Widths -and $c -lt $Widths.Count){$w=$Widths[$c]}
  if($null -eq $w){$m=([string]$Headers[$c]).Length;for($r=1;$r -lt [Math]::Min($nr,201);$r++){$v=$list[$r][$c];$l=$(if($v -is [DateTime]){10}elseif($null -eq $v){0}else{([string]$v).Length});if($l -gt $m){$m=$l}};$w=[Math]::Min(50,[Math]::Max(8,$m+2))}
  $ws.Column($c+1).Width=[double]$w
 }
 if($nr -gt 1){$ws.Cells[1,1,$nr,$nc].AutoFilter=$true};$ws.View.FreezePanes(2,1)
 if($BoldLastRow -and $nr -gt 1){$ws.Cells[$nr,1,$nr,$nc].Style.Font.Bold=$true}
 return $ws
}
function Get-CmbSheetRows([object[]]$Arrays,[int[]]$Order,[string]$Booleans,[object[]]$Lead=$null){
 # Rows for a sheet straight from the column arrays (one row per patient in output order); true/false follow the config's booleans.
 $nc=$Arrays.Count;$off=0;if($Lead){$off=1};$list=[Collections.Generic.List[object[]]]::new([Math]::Max(1,$Order.Length))
 foreach($q in $Order){
  $row=[object[]]::new($nc+$off);if($off){$row[0]=$Lead[$q]}
  for($c=0;$c -lt $nc;$c++){$v=$Arrays[$c][$q];if($null -ne $v){if($v -is [bool]){$v=Get-CmbCellValue $v $Booleans}elseif($v -is [string] -and $v.Length -gt 32000){$v=$v.Substring(0,32000)}};$row[$c+$off]=$v}
  $list.Add($row)
 }
 return ,$list
}
function Write-CmbWorkbook($Result,[string]$Path,$Run){
 $plan=$Result.plan;$b=$plan.config.booleans;$FV=$Result.F;$order=$Result.order;$used=@{}
 $pkg=New-Object OfficeOpenXml.ExcelPackage
 try{
  $cols=@($Result.outCols)
  $rows=Get-CmbSheetRows ([object[]]@(foreach($c in $cols){,$FV[$c.key]})) $order $b
  $null=Add-CmbSheet $pkg $used $plan.config.sheet ([string[]]@($cols|ForEach-Object{$_.header})) $rows ([string[]]@($cols|ForEach-Object{$_.format})) ([object[]]@($cols|ForEach-Object{$_.width}))
  foreach($v in $Result.views){
   $vr=@(foreach($r in $v.rows){,([object[]]@(foreach($x in $r){Get-CmbCellValue $x $b}))});$bold=0;if($v.total){$vr+=,([object[]]@(foreach($x in $v.total){Get-CmbCellValue $x $b}));$bold=1}
   $fm=@(for($i=0;$i -lt $v.groupCount;$i++){''})+@($v.formats)
   $null=Add-CmbSheet $pkg $used $v.name ([string[]]$v.headers) $vr ([string[]]$fm) $null $bold
  }
  $null=Add-CmbSheet $pkg $used 'Audit - Rules' @('Row','Rule','What it does','Effect in this run','Problems','Note') @(Get-CmbEffectRows $Result) $null @(6,10,70,60,40,50)
  $null=Add-CmbSheet $pkg $used 'Audit - Sources' @('Source','Label','Role','File','Sheet','Header row','SHA-256','Rows read','Empty rows','Removed (filters, unique)','Rows without a key','Rows used','Patients') @(foreach($s in $Result.sources){,([object[]]@($s.source,$s.label,$s.role,$s.file,$s.sheet,$s.headerRow,([string]$s.sha256).Substring(0,[Math]::Min(12,([string]$s.sha256).Length)),[double]$s.rowsRead,[double]$s.emptyRows,[double]$s.filtered,[double]$s.noKey,[double]$s.rowsUsed,[double]$s.patients))})
  if($Result.conflicts.Count){$null=Add-CmbSheet $pkg $used 'Audit - Conflicts' @('Patient','Field','Chosen value','Chosen from','Other values [source]','How','Rule row') @(foreach($c in $Result.conflicts){,([object[]]@($c.patient,$c.field,$c.chosen,$c.from,$c.others,$c.how,$(if($c.ruleIndex -ge 0){[double]($c.ruleIndex+1)}else{'(default)'})))}) $null @(18,18,28,14,60,12,9)}
  if($Result.keyConflicts.Count){$null=Add-CmbSheet $pkg $used 'Audit - Key conflicts' @('Key','Shared value','Earlier key','First patient has','Second patient has','Sources') @(foreach($c in $Result.keyConflicts){,([object[]]@($c.key,$c.value,$c.earlierKey,$c.first,$c.second,$c.sources))})}
  if($Result.excluded.Count){$null=Add-CmbSheet $pkg $used 'Audit - Excluded' @('Patient','Removed by rule row','Filter') @(foreach($e in $Result.excluded){$q=$e[0];$ri=$e[1];,([object[]]@([CmbRt]::Text($Result.keyText[$q]),[double]($ri+1),(Get-CmbRuleSummary $plan.rules[$ri])))}) $null @(20,10,80)}
  if($plan.config.lineage){
   $n=$Result.P;$fixed=@{};foreach($t in @('(built-in)','(derived)')){$a=[object[]]::new($n);for($q=0;$q -lt $n;$q++){$a[$q]=$t};$fixed[$t]=$a}
   $arrs=[object[]]@(foreach($c in $cols){if($Result.LIN.ContainsKey($c.key)){,([object[]]$Result.LIN[$c.key])}elseif($c.key.StartsWith('_')){,$fixed['(built-in)']}else{,$fixed['(derived)']}})
   $keys=[object[]]::new($n);for($q=0;$q -lt $n;$q++){$keys[$q]=[CmbRt]::Text($Result.keyText[$q])}
   $null=Add-CmbSheet $pkg $used 'Audit - Value sources' ([string[]](@('Patient')+@($cols|ForEach-Object{$_.header}))) (Get-CmbSheetRows $arrs $order $b $keys)
  }
  $meta=@(,([object[]]@('Config',[string]$Run.configName)),([object[]]@('Config file',[string]$Run.configPath)),([object[]]@('Config SHA-256',[string]$Run.configSha256)),([object[]]@('Built',(Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))),([object[]]@('today() is',$plan.config.asOf.ToString('yyyy-MM-dd'))),([object[]]@('','')),([object[]]@('Rules (exactly as used, one per row)','')))
  $ri=0;foreach($r in $plan.rules){$ri++;$meta+=,([object[]]@([string]$ri,(ConvertTo-CmbRuleText $r)))}
  $null=Add-CmbSheet $pkg $used 'Audit - Config' @('Item','Value') $meta $null @(34,150)
  $tmp=$Path+'.'+$PID+'.tmp';[IO.File]::WriteAllBytes($tmp,$pkg.GetAsByteArray());Move-Item -LiteralPath $tmp -Destination $Path -Force
 }finally{$pkg.Dispose()}
}
function Get-CmbEffectRows($Result){
 $plan=$Result.plan;$out=@()
 for($i=0;$i -lt $plan.rules.Count;$i++){
  $r=$plan.rules[$i];$type='';if($r -is [Collections.IDictionary]){$type=([string](Get-CmbValue $r 'rule' '')).ToLowerInvariant()}
  $e='';if($Result.effects.ContainsKey($i)){$e=[string]$Result.effects[$i]}
  $pr=@($plan.problems|Where-Object{$_.rule -eq ($i+1)}|ForEach-Object{$_.severity+': '+$_.message}) -join ' | '
  $out+=,([object[]]@([double]($i+1),$type,(Get-CmbRuleSummary $r),$e,$pr,$(if($r -is [Collections.IDictionary]){Get-CmbText $r 'note'}else{''})))
 }
 # Mapped fields without a field rule are combined by priority; their effects are listed after the rules.
 foreach($k in @($Result.fieldNotes.Keys)){$f=$plan.fields[$k];$out+=,([object[]]@('','field (default)',($f.name+' = first non-blank value trying '+(@($f.from) -join ', ')+' in order'),[string]$Result.fieldNotes[$k],'','No field rule; combined by source priority'))}
 return $out
}
function Get-CmbRunPath([string]$Id){if($Id -notmatch '^[a-f0-9]{32}$'){throw 'Invalid run ID.'};return (Join-Path $script:Paths.State ('cmbrun-'+$Id+'.json'))}
function Get-CmbRuns{return @(Get-ChildItem $script:Paths.State -Filter 'cmbrun-*.json' -File -ErrorAction SilentlyContinue|ForEach-Object{Json $_.FullName}|Where-Object{$null -ne $_}|Sort-Object -Property @{Expression={ConvertTo-IsoText (Get-P $_ 'requestedUtc' '')}} -Descending)}
function Get-CmbPicks($Plan,[object[]]$Assignments){
 # Source key -> @{pick;file} from a run's saved assignments; the file must still be loaded and unchanged.
 $picks=@{}
 foreach($a in $Assignments){
  $k=([string](Get-P $a 'source' '')).ToLowerInvariant();$tok=[string](Get-P $a 'token' '');if(!$tok){continue}
  $f=Get-CmbFile $tok;if([string](Get-P $f 'sha256' '') -ne [string](Get-P $a 'sha256' '')){throw ('"'+[string]$f.originalName+'" changed since the run was requested.')}
  $sheet=[string](Get-P $a 'sheet' '');$s=@(@(Get-P $f 'sheets' @())|Where-Object{[string]$_.name -eq $sheet});if(!$s.Count){throw ('Sheet "'+$sheet+'" is not in '+[string]$f.originalName+'.')}
  $picks[$k]=@{pick=@{token=$tok;sheet=$sheet;headerRow=[int]$s[0].headerRow;headers=@(Get-P $s[0] 'headers' @());file=[string]$f.originalName};file=$f}
 }
 return $picks
}
function Invoke-CmbRun([string]$Id){
 # Runs in its own worker process: checks the saved config against the chosen files, combines them and writes the workbook.
 $path=Get-CmbRunPath $Id;$run=Json $path;if(!$run){throw 'Run was not found.'}
 $sw=[Diagnostics.Stopwatch]::StartNew()
 try{
  Set-P $run 'workerPid' $PID;Set-P $run 'startedUtc' ([DateTime]::UtcNow.ToString('o'))
  if([string](Get-P $run 'kind' '') -eq 'parity'){
   $pr=Invoke-CmbParity $run $path;$run=Json $path
   Set-P $run 'result' $pr.result;Set-P $run 'parity' $pr.parity;Set-P $run 'outputPath' $pr.outputPath;Set-P $run 'parityPath' $pr.parityPath;Set-P $run 'assignments' @(Get-P $run 'assignments' @())
   $run.state='Completed';$run.percent=100;$run.stage='Compared';Set-P $run 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $run 'seconds' ([Math]::Round($sw.Elapsed.TotalSeconds,1));Save-JsonAtomic $path $run
   Log 'COMBINE_PARITY_COMPLETED' 'OK' ([string]$run.configName+'; '+$pr.parity.matched+' paired, '+$pr.parity.differentRows+' with differences, '+$pr.parity.onlyReports+' only in reports, '+$pr.parity.onlyCombine+' only in Combine');return
  }
  Set-RunProgress $run $path 2 'Checking the config'
  $cfg=Read-CmbConfigText ([string]$run.configText);if(@($cfg.errors).Count){throw ('The config could not be read: '+$cfg.errors[0].message)}
  $rules=@($cfg.rules);$plan0=Get-CmbPlan $rules
  $picks=Get-CmbPicks $plan0 @(Get-P $run 'assignments' @())
  $tables=@{};foreach($k in @($picks.Keys)){$p=$picks[$k].pick;$tables[$k]=@{headers=$p.headers;file=('"'+$p.file+'"')}}
  $plan=Get-CmbPlan $rules $tables -Strict
  if($plan.errorCount){$first=@($plan.problems|Where-Object{$_.severity -eq 'error'}|Select-Object -First 3|ForEach-Object{$(if($_.rule){'row '+$_.rule+': '}else{''})+$_.message});throw ('The config has '+$plan.errorCount+' problem'+$(if($plan.errorCount -eq 1){''}else{'s'})+' to fix first: '+($first -join ' | '))}
  $progress={param([int]$pc,[string]$st)Set-RunProgress $run $path $pc $st}
  $res=Invoke-CmbEngine $plan $picks $progress
  Set-RunProgress $run $path 90 'Writing the workbook'
  $stamp=(Get-Date).ToString('yyyy-MM-dd HHmmss');$base=(Get-SafeFileName ([string]$run.configName))+' - '+$stamp;$out=Join-Path $script:Paths.CombineOutputs ($base+'.xlsx');$n=2
  while(Test-Path -LiteralPath $out){$out=Join-Path $script:Paths.CombineOutputs ($base+' ('+$n+').xlsx');$n++}
  Write-CmbWorkbook $res $out $run
  Set-RunProgress $run $path 97 'Saving the summary';$run=Json $path
  Set-P $run 'result' (Get-CmbRunSummary $res)
  Set-P $run 'outputPath' $out;$run.state='Completed';$run.percent=100;$run.stage='Combined';Set-P $run 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $run 'seconds' ([Math]::Round($sw.Elapsed.TotalSeconds,1));Save-JsonAtomic $path $run
  Log 'COMBINE_COMPLETED' 'OK' ([string]$run.configName+'; '+$res.order.Length+' patients in '+[Math]::Round($sw.Elapsed.TotalSeconds,1)+'s')
 }catch{$msg=$_.Exception.Message;$cur=Json $path;if(!$cur){$cur=$run};$cur.state='Failed';$cur.stage='Run failed';Set-P $cur 'errorSummary' $msg;Set-P $cur 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $cur;Log 'COMBINE_FAILED' 'FAILED' ($msg+' | '+$_.ScriptStackTrace)}
}
function Get-CmbRunSummary($Res){
 # What the Output page shows without opening the workbook: counts, each rule's effect, sources, and the first rows of every table.
 $plan=$Res.plan;$b=$plan.config.booleans;$cols=@($Res.outCols);$n=[Math]::Min($script:CmbPreviewRows,$Res.order.Length)
 $prev=@(for($i=0;$i -lt $n;$i++){$q=$Res.order[$i];,([string[]]@(foreach($c in $cols){Format-CmbDisplay $Res.F[$c.key][$q] $b $c.format}))})
 $views=@(foreach($v in $Res.views){
  $fm=@(for($i=0;$i -lt $v.groupCount;$i++){''})+@($v.formats)
  $fmtRow={param($r)[string[]]@(for($i=0;$i -lt $r.Count;$i++){Format-CmbDisplay $r[$i] $b ([string]$fm[$i])})}
  [ordered]@{name=$v.name;headers=@($v.headers);groupCount=$v.groupCount;rows=@($v.rows|Select-Object -First 200|ForEach-Object{,(& $fmtRow $_)});totalRows=$v.rows.Count;total=$(if($v.total){,(& $fmtRow $v.total)}else{$null});patients=$v.patients}
 })
 $effects=@(Get-CmbEffectRows $Res|ForEach-Object{[ordered]@{index=$(if([string]$_[0]){[int]$_[0]}else{0});type=[string]$_[1];summary=[string]$_[2];effect=[string]$_[3];problems=[string]$_[4];note=[string]$_[5]}})
 return [ordered]@{patients=$Res.order.Length;matched=$Res.P;records=$Res.records;noKey=$Res.noKey;conflicts=$Res.conflictCount;conflictRows=$Res.conflicts.Count;keyConflicts=$Res.keyConflicts.Count;excluded=$Res.excluded.Count;engineSeconds=$Res.seconds;asOf=$plan.config.asOf.ToString('yyyy-MM-dd');sheet=$plan.config.sheet
  headers=@($cols|ForEach-Object{$_.header});preview=$prev;views=$views;effects=$effects;sources=@($Res.sources|ForEach-Object{$x=[ordered]@{};foreach($k in @($_.Keys)){if($k -ne 'token'){$x[$k]=$_[$k]}};$x})
  conflictSample=@($Res.conflicts|Select-Object -First 50)}
}
function Start-CmbWorker([string]$Id){return (Start-RunWorker 'RunCombineId' $Id ('combine-'+$Id))}
function Update-CmbLiveness($Run){
 $id=[string]$Run.runId;$reason=Get-RunStallReason $Run $id 'combine' 'Run it again.' ('combine-'+$id)
 if(!$reason){return $false}
 $Run.state='Failed';$Run.stage='Run failed';Set-P $Run 'errorSummary' $reason;Save-JsonAtomic (Get-CmbRunPath $id) $Run;Log 'COMBINE_INTERRUPTED' 'WARN' $id;return $true
}
function Get-CmbRunModel([string]$Id,[switch]$Summary){
 $run=Json (Get-CmbRunPath $Id);if(!$run){throw 'That run was not found.'}
 $null=Update-CmbLiveness $run
 $m=[ordered]@{runId=[string]$run.runId;state=[string]$run.state;percent=[int](Get-P $run 'percent' 0);stage=[string](Get-P $run 'stage' '');requestedUtc=(ConvertTo-IsoText (Get-P $run 'requestedUtc' ''));completedUtc=(ConvertTo-IsoText (Get-P $run 'completedUtc' ''));seconds=(Get-P $run 'seconds' $null);errorSummary=[string](Get-P $run 'errorSummary' '');configName=[string](Get-P $run 'configName' '');assignments=@(Get-P $run 'assignments' @()|ForEach-Object{[ordered]@{source=[string]$_.source;file=[string]$_.file;sheet=[string]$_.sheet;rows=[int](Get-P $_ 'rows' 0)}});result=$null;outputUrl=$null;outputFile='';outputError='';folder=$script:Paths.CombineOutputs;patients=$null;kind=[string](Get-P $run 'kind' 'combine');parity=$null;parityUrl=$null;parityFile='';differentRows=$null}
 if(!$m.kind){$m.kind='combine'}
 if($m.state -eq 'Completed'){
  $r=Get-P $run 'result' $null;if($r){$m.patients=[int](Get-P $r 'patients' 0);if(!$Summary){$m.result=$r}}
  $op=[string](Get-P $run 'outputPath' '');if($op -and (Test-Path -LiteralPath $op)){$m.outputUrl='/combine/output.xlsx?id='+$m.runId;$m.outputFile=$op}else{$m.outputError='The workbook was moved or deleted from the combine\outputs folder.'}
  if($m.kind -eq 'parity'){$pr=Get-P $run 'parity' $null;if($pr){$m.parity=$pr;$m.differentRows=[int](Get-P $pr 'differentRows' 0)};$pp=[string](Get-P $run 'parityPath' '');if($pp -and (Test-Path -LiteralPath $pp)){$m.parityUrl='/combine/parity.xlsx?id='+$m.runId;$m.parityFile=$pp}}
 }
 return $m
}
function New-CmbRun($Body){
 $name=[string](Get-P $Body 'name' '');$c=Read-CmbConfigFile $name;if(@($c.errors).Count){throw ('The config file could not be read (line '+$c.errors[0].line+': '+$c.errors[0].message+'); fix it on the Rules page.')}
 $chosen=@{};$ch=Get-P $Body 'choices' $null;if($ch){foreach($p in $ch.PSObject.Properties){$chosen[$p.Name.ToLowerInvariant()]=[string]$p.Value}}
 $files=@(Get-CmbFiles|Where-Object{Test-Path -LiteralPath ([string](Get-P $_ 'path' ''))})
 $plan0=Get-CmbPlan $c.rules;$assign=Get-CmbAssignments $plan0 $files $chosen;$plan=Get-CmbPlan $c.rules (Get-CmbTablesFromAssignments $assign) -Strict
 if($plan.errorCount){$first=@($plan.problems|Where-Object{$_.severity -eq 'error'}|Select-Object -First 3|ForEach-Object{$(if($_.rule){'row '+$_.rule+': '}else{''})+$_.message});throw ('Fix the config first ('+$plan.errorCount+' problem'+$(if($plan.errorCount -eq 1){''}else{'s'})+'): '+($first -join ' | '))}
 foreach($r in @(Get-CmbRuns)){if([string](Get-P $r 'state' '') -in @('Starting','Running') -and !(Update-CmbLiveness $r)){throw 'A combine run is already running; wait for it to finish.'}}
 $asg=@(foreach($k in @($assign.Keys)){$p=$assign[$k].pick;if(!$p){continue};$f=Get-CmbFile $p.token;[ordered]@{source=$assign[$k].source;token=$p.token;sheet=$p.sheet;file=$p.file;sha256=[string]$f.sha256;rows=$p.rows}})
 $sha=[Security.Cryptography.SHA256]::Create();try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($c.text))) -replace '-','')}finally{$sha.Dispose()}
 $id=[Guid]::NewGuid().ToString('N');$path=Get-CmbRunPath $id
 $run=[ordered]@{cmbRunVersion=1;runId=$id;state='Starting';percent=1;stage='Launching the combine worker';requestedUtc=[DateTime]::UtcNow.ToString('o');configName=$c.name;configPath=$c.path;configText=$c.text;configSha256=$hash;assignments=$asg;workerPid=0}
 Save-JsonAtomic $path $run;Log 'COMBINE_REQUESTED' 'OK' ($c.name+'; '+(@($asg|ForEach-Object{$_.source+'='+$_.file}) -join ', '))
 if($script:CmbInline){Invoke-CmbRun $id}
 else{try{$null=Start-CmbWorker $id}catch{$run.state='Failed';$run.stage='Worker launch failed';$run['errorSummary']=$_.Exception.Message;Save-JsonAtomic $path $run;Log 'COMBINE_FAILED' 'FAILED' $_.Exception.Message}}
 foreach($old in @(Get-CmbRuns|Select-Object -Skip $script:CmbRunKeep)){$oid=[string](Get-P $old 'runId' '');if($oid -match '^[a-f0-9]{32}$'){Remove-Item -LiteralPath (Get-CmbRunPath $oid) -Force -ErrorAction SilentlyContinue;Remove-Item -LiteralPath (Join-Path $script:Paths.Logs ('combine-'+$oid+'.out.log')),(Join-Path $script:Paths.Logs ('combine-'+$oid+'.err.log')) -Force -ErrorAction SilentlyContinue}}
 return (Get-CmbRunModel $id)
}
function Send-CmbOutput($Context,[string]$Id,[string]$Which='outputPath'){
 $run=Json (Get-CmbRunPath $Id);if(!$run -or [string](Get-P $run 'state' '') -ne 'Completed'){throw 'That run is not complete.'}
 $full=[IO.Path]::GetFullPath([string](Get-P $run $Which ''));$root=[IO.Path]::GetFullPath($script:Paths.CombineOutputs).TrimEnd([IO.Path]::DirectorySeparatorChar)+[IO.Path]::DirectorySeparatorChar
 if(!$full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or !(Test-Path -LiteralPath $full)){throw 'The workbook is not available; it may have been moved or deleted from the combine\outputs folder.'}
 $name=[IO.Path]::GetFileName($full) -replace '[^\x20-\x7E]','_' -replace '"',''
 Send-Bytes $Context 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ([IO.File]::ReadAllBytes($full)) ('attachment; filename="'+$name+'"')
}

# --- Combine LLM round trip: the request copied to the clipboard is one JSON document laid out one item per line, so it pastes into Excel
# one row per line (every line starts with a quote or a bracket, never with = + - @) and an LLM reading the sheet sees ordinary JSON.
function Format-CmbPromptDocument($Doc){
 $lines=[Collections.Generic.List[string]]::new();$lines.Add('{');$keys=@($Doc.Keys)
 for($i=0;$i -lt $keys.Count;$i++){
  $k=$keys[$i];$v=$Doc[$k];$comma=$(if($i -lt $keys.Count-1){','}else{''});$name=(ConvertTo-CmbJsonString $k)+': '
  if($v -is [Array]){
   if($v.Count -eq 0){$lines.Add($name+'[]'+$comma);continue}
   $lines.Add($name+'[');for($j=0;$j -lt $v.Count;$j++){$lines.Add((ConvertTo-CmbJson $v[$j])+$(if($j -lt $v.Count-1){','}else{''}))};$lines.Add(']'+$comma)
  }elseif($v -is [Collections.IDictionary]){
   $lines.Add($name+'{');$ik=@($v.Keys);for($j=0;$j -lt $ik.Count;$j++){$lines.Add((ConvertTo-CmbJsonString ([string]$ik[$j]))+': '+(ConvertTo-CmbJson $v[$ik[$j]])+$(if($j -lt $ik.Count-1){','}else{''}))};$lines.Add('}'+$comma)
  }else{$lines.Add($name+(ConvertTo-CmbJson $v)+$comma)}
 }
 $lines.Add('}');return ($lines -join "`r`n")
}
function Get-CmbPrompt([string]$Name,[string]$Text,[string]$Request,[string[]]$Tokens,[bool]$Examples){
 # The request for the enterprise LLM: what to do, how the system works, the current config with its problems, and the files' fingerprints.
 $read=Read-CmbConfigText $Text;$rules=@($read.rules)
 $files=@(foreach($t in $Tokens){if($t -match '^[a-f0-9]{32}$'){try{Get-CmbFile $t}catch{}}})
 $check=Get-CmbCheck $rules $files
 $usedBy=@{};foreach($s in $check.sources){if($s.chosen){$usedBy[$s.chosen]=$s.source}}
 $spec=@(foreach($t in @($script:CmbRuleSpec.Keys)){$s=$script:CmbRuleSpec[$t];$ex=$null;try{$ex=ConvertFrom-CmbJson $s.example}catch{};[ordered]@{rule=$t;purpose=$s.purpose;properties=$s.props;example=$ex}})
 $task=$Request.Trim();if(!$task){$task=$(if($rules.Count){'Review the current config against the files described below. Fix anything that is wrong, missing or would lose data, and keep everything that is already right.'}else{'Write a config that combines the files described below into one row per patient, with summary views by provider, location, risk pool and region where the data allows.'})}
 $doc=[ordered]@{}
 $doc['purpose']='This is a request to write a Data Combiner config. Read every part, then answer exactly as answerFormat says. Everything you need is here; the program that runs the config reads only the rule types, properties, cleaning steps and functions described below.'
 $doc['task']=$task
 $doc['howItWorks']=@($script:CmbGuide)
 $doc['ruleTypes']=$spec
 $doc['howMethods']=$script:CmbHowText
 $doc['cleaningSteps']=$script:CmbCleanText
 $doc['expressions']=@($script:CmbExprGuide)
 $doc['functions']=$script:CmbFunctionText
 $doc['answerFormat']=@($script:CmbAnswerGuide)
 $doc['configName']=$Name
 $doc['currentConfig']=@($rules)
 $doc['problemsInCurrentConfig']=@(foreach($r in $check.rules){foreach($p in $r.problems){[ordered]@{row=$r.index;severity=$p.severity;message=$p.message}}})+@(foreach($p in $check.general){[ordered]@{row=0;severity=$p.severity;message=$p.message}})
 $doc['files']=@(foreach($f in $files){foreach($s in @(Get-P $f 'sheets' @())){$k=[string]$f.token+'|'+[string]$s.name;$o=[ordered]@{file=[string]$f.originalName;sheet=[string]$s.name;headerRow=[int]$s.headerRow;rows=[int]$s.rows;layoutId=[string]$s.layoutId};if($usedBy.ContainsKey($k)){$o['usedBySource']=$usedBy[$k]};$o}})
 $doc['columns']=@(foreach($f in $files){foreach($s in @(Get-P $f 'sheets' @())){foreach($c in @(Get-P $s 'columns' @())){
  $o=[ordered]@{file=[string]$f.originalName;sheet=[string]$s.name;column=[string]$c.letter;header=[string]$c.header;type=[string]$c.type;filled=[int]$c.filled;blank=[int]$c.blank;distinct=[int]$c.distinct}
  if([bool]$c.unique){$o['unique']=$true}
  $sh=@(@(Get-P $c 'patterns' @())|ForEach-Object{[string]$_.shape+' '+[Math]::Round(100*[double]$_.share)+'%'});if($sh.Count){$o['shapes']=$sh}
  $cm=@(@(Get-P $c 'common' @())|ForEach-Object{[string]$_.value+' ('+[int]$_.rows+')'});if($cm.Count){$o['values']=$cm}
  if([string](Get-P $c 'min' '')){$o['min']=[string]$c.min;$o['max']=[string]$c.max}
  if($Examples){$ex=@(Get-P $c 'examples' @());if($ex.Count){$o['examples']=@($ex|ForEach-Object{[string]$_})}}
  $o
 }}})
 $out=Format-CmbPromptDocument $doc
 return [ordered]@{text=$out;lines=($out -split "`n").Count;characters=$out.Length;files=$files.Count;rules=$rules.Count}
}
function ConvertTo-CmbFieldName([string]$Header){
 $w=@(($Header -replace '[^A-Za-z0-9]+',' ').Trim() -split ' '|Where-Object{$_});if($w.Count -eq 0){return 'Field'}
 $s=(@($w|ForEach-Object{$_.Substring(0,1).ToUpperInvariant()+$_.Substring(1)}) -join '');if($s -notmatch '^[A-Za-z]'){$s='F'+$s};if($s.Length -gt 40){$s=$s.Substring(0,40)};return $s
}
function New-CmbDraft([string]$Name,[string[]]$Tokens){
 # A first config from loaded files: one source per file (matched by its most distinctive headers), every column mapped (the same header in
 # several files becomes one field), the patient id as key when one is recognizable, explicit priority for shared fields, every field output,
 # and views for provider, location, risk pool and region columns. Rules a person must check carry a QUESTION note.
 $files=@(foreach($t in $Tokens){Get-CmbFile $t});if($files.Count -eq 0){throw 'Choose at least one loaded file.'}
 $tables=@(foreach($f in $files){$s=@(@(Get-P $f 'sheets' @())|Sort-Object @{Expression={[int]$_.rows}} -Descending)[0];@{file=$f;sheet=$s}})
 $count=@{};foreach($t in $tables){$seen=@{};foreach($h in @(Get-P $t.sheet 'headers' @())){$k=ConvertTo-CmbHeaderKey ([string]$h);if(!$seen.ContainsKey($k)){$seen[$k]=$true;$count[$k]=1+[int]$count[$k]}}}
 $fieldOf=@{};$fieldUsed=@{};$ids=@{};$rules=[Collections.Generic.List[object]]::new()
 $cfg=[ordered]@{};$cfg['rule']='config';$cfg['name']=$Name;$cfg['description']=('Draft made from '+$files.Count+' file'+$(if($files.Count -eq 1){''}else{'s'})+' on '+(Get-Date).ToString('yyyy-MM-dd')+'; review every rule.');$cfg['note']='Draft generated by the program; check the key, the field priorities and the views.';$rules.Add($cfg)
 $keyRx='\b(mrn|medical record( number| no| num| #)?|patient id|pat id|patient number|pat num|enterprise id|empi|member id|person id|chart number)\b'
 $keyField=$null;$keyVotes=@{}
 $srcRules=[Collections.Generic.List[object]]::new();$mapRules=[Collections.Generic.List[object]]::new();$srcIds=@()
 foreach($t in $tables){
  $base=([IO.Path]::GetFileNameWithoutExtension([string]$t.file.originalName)).ToLowerInvariant() -replace '[^a-z0-9]+','_' -replace '^_+|_+$','';if($base -notmatch '^[a-z]'){$base='s_'+$base};if($base.Length -gt 24){$base=$base.Substring(0,24).TrimEnd('_')}
  $id=$base;$k=2;while($ids.ContainsKey($id)){$id=$base+'_'+$k;$k++};$ids[$id]=$true;$srcIds+=$id
  $cols=@(Get-P $t.sheet 'columns' @());$headers=@(Get-P $t.sheet 'headers' @()|ForEach-Object{[string]$_})
  $distinctive=@($headers|Where-Object{$count[(ConvertTo-CmbHeaderKey $_)] -eq 1});$keyHeaders=@($headers|Where-Object{(ConvertTo-CmbHeaderKey $_) -match $keyRx})
  $match=@(@($keyHeaders|Select-Object -First 1)+@($distinctive|Select-Object -First 5)+@($headers|Select-Object -First 6)|Select-Object -Unique|Select-Object -First 6)
  $s=[ordered]@{};$s['rule']='source';$s['source']=$id;$s['label']=[string]$t.file.originalName;$s['match']=$match;if([IO.Path]::GetExtension([string]$t.file.originalName) -ieq '.xlsx'){$s['sheet']=[string]$t.sheet.name}
  $s['note']=('Rows: '+[int]$t.sheet.rows+'. Matched by '+$(if($distinctive.Count){'columns only this file has'}else{'its columns (no column is unique to it; add fileName if two files look alike)'})+'.');$srcRules.Add($s)
  foreach($c in $cols){
   $h=[string]$c.header;$hk=ConvertTo-CmbHeaderKey $h;if([int]$c.filled -eq 0){continue}
   if(!$fieldOf.ContainsKey($hk)){$fn=ConvertTo-CmbFieldName $h;$b2=$fn;$k=2;while($fieldUsed.ContainsKey($fn.ToLowerInvariant())){$fn=$b2+$k;$k++};$fieldUsed[$fn.ToLowerInvariant()]=$true;$fieldOf[$hk]=$fn}
   $fn=$fieldOf[$hk];$m=[ordered]@{};$m['rule']='map';$m['source']=$id;$m['field']=$fn;$m['column']=$h
   $type=[string]$c.type;$isKey=($hk -match $keyRx)
   if($isKey){$keyVotes[$fn]=1+[int]$keyVotes[$fn];$shapes=@(@(Get-P $c 'patterns' @())|ForEach-Object{[string]$_.shape});if(@($shapes|Where-Object{$_ -notmatch '^9+$'}).Count -eq 0){$m['clean']=@('digits','stripLeadingZeros');$m['note']='Patient id: digits only, without leading zeros, so it matches across files'}else{$m['clean']=@('upper');$m['note']='Patient id'}}
   elseif($type -eq 'date'){$m['type']='date'}elseif($type -eq 'number'){$m['type']='number'}elseif($type -eq 'bool'){$m['type']='bool'}
   elseif($type -like 'mixed*'){$m['note']=('QUESTION: values are '+$type+'; choose a type or cleaning')}
   $mapRules.Add($m)
  }
 }
 foreach($r in $srcRules){$rules.Add($r)};foreach($r in $mapRules){$rules.Add($r)}
 if($keyVotes.Count){$keyField=@($keyVotes.Keys|Sort-Object @{Expression={$keyVotes[$_]}} -Descending)[0]}
 $kr=[ordered]@{};$kr['rule']='key'
 if($keyField){$kr['fields']=@($keyField);$missing=@(for($i=0;$i -lt $tables.Count;$i++){$has=$false;foreach($m in $mapRules){if($m['source'] -eq $srcIds[$i] -and $m['field'] -eq $keyField){$has=$true}};if(!$has){$srcIds[$i]}});$kr['note']=$(if($missing.Count){'QUESTION: '+($missing -join ', ')+' has no '+$keyField+' column; map one, or add a second key rule (for example name and date of birth)'}else{$keyField+' is in every file'})}
 else{$kr['fields']=@('MRN');$kr['note']='QUESTION: no patient id column was recognized; set fields to the column that identifies a patient in every file'}
 $rules.Add($kr)
 $shared=@{};foreach($m in $mapRules){$f=[string]$m['field'];if(!$shared.ContainsKey($f)){$shared[$f]=@()};$shared[$f]+=[string]$m['source']}
 foreach($f in @($fieldOf.Values|Select-Object -Unique)){if($f -eq $keyField){continue};$s=@($shared[$f]|Select-Object -Unique);if($s.Count -lt 2){continue};$fr=[ordered]@{};$fr['rule']='field';$fr['field']=$f;$fr['how']='priority';$fr['from']=$s;$fr['note']=('In '+$s.Count+' files; the first non-blank value wins in this order. QUESTION: is this the right order, or should the latest value win?');$rules.Add($fr)}
 $outFields=@();if($keyField){$outFields+=$keyField};foreach($m in $mapRules){$f=[string]$m['field'];if($outFields -notcontains $f){$outFields+=$f}};$outFields+='_sources'
 foreach($f in $outFields){$o=[ordered]@{};$o['rule']='output';$o['field']=$f;if($f -eq '_sources'){$o['header']='Sources';$o['note']='Which files the patient was found in'};$rules.Add($o)}
 $viewDefs=@(@('By provider','\b(pcp|provider|physician|attributed|clinician|doctor)\b'),@('By location','\b(location|site|practice|clinic|department|office)\b'),@('By risk pool','\b(risk pool|pool|payer|payor|plan|contract)\b'),@('By region','\b(region|market|division|area)\b'))
 foreach($vd in $viewDefs){
  $hit=@($fieldOf.Keys|Where-Object{$_ -match $vd[1] -and $_ -notmatch '\b(npi|id|date|phone|address)\b'}|Sort-Object @{Expression={$count[$_]}} -Descending);if($hit.Count -eq 0){continue}
  $f=$fieldOf[$hit[0]];$v=[ordered]@{};$v['rule']='view';$v['view']=$vd[0];$v['groupBy']=@($f);$v['sortBy']='Patients';$v['desc']=$true;$v['note']=('QUESTION: grouped by '+$f+'; check this is the right column');$rules.Add($v)
  $m=[ordered]@{};$m['rule']='measure';$m['view']=$vd[0];$m['name']='Patients';$m['agg']='count';$rules.Add($m)
 }
 return ,@($rules)
}
# --- Combine side by side: the reports' patient facts (New-PatientFacts) written as a config, and a check that runs both on the current
# Provider Analysis files and compares them provider by provider, patient by patient, field by field.
$script:CmbPaConfigName='Provider Analysis facts'
$script:CmbProviderAnalysisConfig=@'
[
{"rule":"config","name":"Provider Analysis facts","description":"The patient facts behind the provider reports, as rules: one row per provider and patient, joined the way the report code (New-PatientFacts) joins the seven sources.","note":"Generated from the report code. The side-by-side check on the Output page compares this config with the reports, patient by patient."},
{"rule":"source","source":"export","label":"Export","match":["MemberID","Provider Name","Last ACV Date","Total Care Gaps"],"fileName":"Export*","read":"text","note":"Defines the patients: a provider's report lists the Export rows under that provider's name"},
{"rule":"source","source":"quality","label":"PtListQuality","match":["MemberID","Provider Name","OMW Critical Due Date","BCS"],"fileName":"PtListQuality*","read":"text","unique":["NPI","MemberID"],"note":"HEDIS status per measure; a Member ID listed twice under one provider is ignored, as in the reports"},
{"rule":"source","source":"dsc","label":"Diabetes Scorecard","match":["Member ID","Provider","Eye Exam Gap Status","% eGFR last 12 mo."],"fileName":"DiabetesScorecard*","read":"text","unique":["NPI","MemberID"],"required":false,"note":"Diabetic patients; its KED, EED, GSD and Med Adherence DM override PtListQuality"},
{"rule":"source","source":"serial","label":"Serial Scheduling","match":["Member ID","Provider Name","PCP Visit Dates"],"fileName":"SerialScheduling*","read":"text","unique":["NPI","MemberID"],"required":false,"note":"PCP visit dates, A1c date, future visits"},
{"rule":"source","source":"rpo","label":"Risk Population Outreach","match":["Member ID","Epic Pcp","Next Acv"],"fileName":"RiskPopulationOutreach*","read":"text","unique":["NPI","MemberID"],"required":false,"note":"Next ACV (a booked ACV clears ACV due)"},
{"rule":"source","source":"hr_crh","label":"HR-CRH","match":["Patient Insurance ID","PCP Name"],"fileName":"HR-CRH*","read":"text","required":false,"note":"High-risk list; used for a provider whose profile names them in HR-CRH"},
{"rule":"source","source":"hr_riv","label":"HR-RIVPHNYCMM","match":["Patient Insurance ID","PCP Name"],"fileName":"HR-RIVPHNYCMM*","read":"text","required":false,"note":"High-risk list; used only for providers with no HR-CRH name, as in the reports"},
{"rule":"map","source":"export","field":"MemberID","column":"MemberID"},
{"rule":"map","source":"export","field":"ExportProvider","column":"Provider Name"},
{"rule":"map","source":"export","field":"NPI","expr":"lookup('profile:Export', [Provider Name], 'npi')","note":"The provider whose saved profile uses this Export name; rows of providers without a profile get no key"},
{"rule":"map","source":"export","field":"First","column":"First Name"},
{"rule":"map","source":"export","field":"Last","column":"Last Name"},
{"rule":"map","source":"export","field":"DOB","column":"Date of Birth","type":"date"},
{"rule":"map","source":"export","field":"Payer","column":"Payer"},
{"rule":"map","source":"export","field":"LastAcv","column":"Last ACV Date","type":"date"},
{"rule":"map","source":"export","field":"LastQemNonPcp","column":"Last QEM date with non-PCP","type":"date"},
{"rule":"map","source":"export","field":"PcpVisit","column":"Last QEM Visit Date with any PCP in assigned TIN","type":"date"},
{"rule":"map","source":"export","field":"CompletedAttest","column":"Completed Attestations"},
{"rule":"map","source":"export","field":"IncompletedAttest","column":"Incompleted Attestations"},
{"rule":"map","source":"export","field":"OpenICD","column":"Open ICDs"},
{"rule":"map","source":"export","field":"NewPatientFlag","column":"New Patient"},
{"rule":"map","source":"export","field":"TotalCareGaps","column":"Total Care Gaps"},
{"rule":"map","source":"export","field":"TcmFlag","column":"TCM"},
{"rule":"map","source":"export","field":"Phone","column":"Phone Number"},
{"rule":"map","source":"export","field":"ActiveDate","column":"Active Date","type":"date"},
{"rule":"map","source":"quality","field":"MemberID","column":"MemberID"},
{"rule":"map","source":"quality","field":"NPI","expr":"lookup('profile:PtListQuality', [Provider Name], 'npi')"},
{"rule":"map","source":"quality","field":"Q_BCS","column":"BCS"},
{"rule":"map","source":"quality","field":"Q_COLO","column":"COLO"},
{"rule":"map","source":"quality","field":"Q_EED","column":"EED"},
{"rule":"map","source":"quality","field":"Q_GSD","column":"GSD"},
{"rule":"map","source":"quality","field":"Q_CBP","column":"CBP"},
{"rule":"map","source":"quality","field":"Q_OMW","column":"OMW"},
{"rule":"map","source":"quality","field":"Q_KED","column":"KED"},
{"rule":"map","source":"quality","field":"Q_SPC","column":"SPC"},
{"rule":"map","source":"quality","field":"Q_MAD","column":"MAD"},
{"rule":"map","source":"quality","field":"Q_MAC","column":"MAC"},
{"rule":"map","source":"quality","field":"Q_MAH","column":"MAH"},
{"rule":"map","source":"quality","field":"Q_SUPD","column":"SUPD"},
{"rule":"map","source":"quality","field":"Q_COB","column":"COB"},
{"rule":"map","source":"quality","field":"Q_POLY","column":"POLY"},
{"rule":"map","source":"quality","field":"OmwDue","column":"OMW Critical Due Date","type":"date"},
{"rule":"map","source":"dsc","field":"DscProvider","column":"Provider","clean":["fillDown"],"note":"The scorecard names the provider only on the first row of each group"},
{"rule":"map","source":"dsc","field":"NPI","expr":"lookup('profile:DiabetesScorecard', DscProvider, 'npi')"},
{"rule":"map","source":"dsc","field":"MemberID","column":"Member ID"},
{"rule":"map","source":"dsc","field":"MRN","column":"Mrn","clean":["blankIf:0"]},
{"rule":"map","source":"dsc","field":"D_KED","column":"KED"},
{"rule":"map","source":"dsc","field":"D_EED","column":"EED"},
{"rule":"map","source":"dsc","field":"D_GSD","column":"GSD"},
{"rule":"map","source":"dsc","field":"MedAdhDM","column":"Med Adherence DM"},
{"rule":"map","source":"dsc","field":"EgfrPct","column":"% eGFR last 12 mo."},
{"rule":"map","source":"dsc","field":"UacrPct","column":"% uACR last 12 mo."},
{"rule":"map","source":"dsc","field":"NextAppt","column":"Next Appt Date","type":"date"},
{"rule":"map","source":"dsc","field":"NextApptSpecialty","column":"Next Appt Specialty"},
{"rule":"map","source":"dsc","field":"A1cText","column":"Avg Last A1c"},
{"rule":"map","source":"dsc","field":"EyeStatus","column":"Eye Exam Gap Status"},
{"rule":"map","source":"dsc","field":"EyeDate","column":"Eye Exam Date","type":"date"},
{"rule":"map","source":"serial","field":"NPI","expr":"lookup('profile:SerialScheduling', [Provider Name], 'npi')"},
{"rule":"map","source":"serial","field":"MemberID","column":"Member ID"},
{"rule":"map","source":"serial","field":"MRN","column":"MRN","clean":["blankIf:0"]},
{"rule":"map","source":"serial","field":"PcpVisitDates","column":"PCP Visit Dates"},
{"rule":"map","source":"serial","field":"A1cDate","column":"A1c Date","type":"date"},
{"rule":"map","source":"serial","field":"FutureVisits","column":"Future PCP Visits 2026"},
{"rule":"map","source":"rpo","field":"NPI","expr":"lookup('profile:RiskPopulationOutreach', [Epic Pcp], 'npi')"},
{"rule":"map","source":"rpo","field":"MemberID","column":"Member ID"},
{"rule":"map","source":"rpo","field":"MRN","column":"Mrn","clean":["blankIf:0"]},
{"rule":"map","source":"rpo","field":"NextAcv","column":"Next Acv","type":"date"},
{"rule":"map","source":"hr_crh","field":"NPI","expr":"lookup('profile:HR-CRH', [PCP Name], 'npi')"},
{"rule":"map","source":"hr_crh","field":"MemberID","column":"Patient Insurance ID"},
{"rule":"map","source":"hr_crh","field":"HrFirst","column":"Patient First Name"},
{"rule":"map","source":"hr_crh","field":"HrLast","column":"Patient Last Name"},
{"rule":"map","source":"hr_crh","field":"HrDOB","column":"DOB","type":"date"},
{"rule":"map","source":"hr_riv","field":"NPI","expr":"lookup('profile:HR-RIVPHNYCMM', [PCP Name], 'npi')"},
{"rule":"map","source":"hr_riv","field":"MemberID","column":"Patient Insurance ID"},
{"rule":"map","source":"hr_riv","field":"HrFirst","column":"Patient First Name"},
{"rule":"map","source":"hr_riv","field":"HrLast","column":"Patient Last Name"},
{"rule":"map","source":"hr_riv","field":"HrDOB","column":"DOB","type":"date"},
{"rule":"key","fields":["NPI","MemberID"],"sources":["export","quality","dsc","serial","rpo"],"note":"Each provider's report joins only that provider's rows, by Member ID, so a patient here is one provider and one Member ID"},
{"rule":"key","fields":["NPI","MemberID"],"sources":["hr_crh","hr_riv"],"match":"prefix","against":["export"],"note":"High-risk IDs can be cut short: one joins the single Export Member ID that begins with it; several matches leave it out; none makes a high-risk-only patient"},
{"rule":"field","field":"MemberID","from":["export","hr_crh","hr_riv"],"conflicts":"ignore","note":"The Export Member ID; a shortened high-risk ID only for high-risk-only patients"},
{"rule":"field","field":"MRN","from":["dsc","serial","rpo"],"note":"First MRN other than 0 from the scorecard, Serial Scheduling, then Risk Population Outreach"},
{"rule":"field","field":"InExport","how":"exists","from":["export"]},
{"rule":"field","field":"HighRisk","how":"exists","from":["hr_crh","hr_riv"],"note":"On the provider's high-risk list"},
{"rule":"field","field":"Diabetic","how":"exists","from":["dsc"],"note":"On the Diabetes Scorecard"},
{"rule":"derive","field":"Provider","expr":"lookup('profiles', NPI, 'provider')"},
{"rule":"derive","field":"Location","expr":"lookup('profiles', NPI, 'location')"},
{"rule":"derive","field":"RiskPool","expr":"lookup('profiles', NPI, 'riskPool')"},
{"rule":"derive","field":"FirstName","expr":"if(InExport, First, HrFirst)"},
{"rule":"derive","field":"LastName","expr":"if(InExport, Last, HrLast)"},
{"rule":"derive","field":"BirthDate","expr":"if(InExport, DOB, HrDOB)"},
{"rule":"derive","field":"ACOR","expr":"if(InExport, Payer = 'ACOR', blank)"},
{"rule":"derive","field":"ACVDue","expr":"LastAcv is not blank and if(Payer = 'ACOR', daysBetween(LastAcv, today()) >= 366, LastAcv <= date(year(today()) & '-01-01')) and not (NextAcv is not blank and NextAcv >= today())","note":"ACOR: no ACV in the last 366 days; others: none this calendar year (as in the reports, an ACV dated January 1 counts as last year); a future ACV in Risk Population Outreach clears it"},
{"rule":"derive","field":"Attest","expr":"coalesce(number(replace(CompletedAttest, '%$', '')), 0) + coalesce(number(replace(IncompletedAttest, '%$', '')), 0) > 0"},
{"rule":"derive","field":"NewPatient","expr":"upper(trim(NewPatientFlag)) in ('1','Y','YES','TRUE')"},
{"rule":"derive","field":"Egfr","expr":"coalesce(number(replace(EgfrPct, '%$', '')), 0)"},
{"rule":"derive","field":"Uacr","expr":"coalesce(number(replace(UacrPct, '%$', '')), 0)"},
{"rule":"derive","field":"S_BCS","expr":"if(Q_BCS in ('Needed','Non-compliant'), 'open', if(Q_BCS in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_COLO","expr":"if(Q_COLO in ('Needed','Non-compliant'), 'open', if(Q_COLO in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_EED","expr":"if(D_EED in ('Needed','Non-compliant'), 'open', if(D_EED in ('Completed','Compliant'), 'closed', if(Q_EED in ('Needed','Non-compliant'), 'open', if(Q_EED in ('Completed','Compliant'), 'closed', blank))))","note":"Scorecard EED overrides PtListQuality"},
{"rule":"derive","field":"S_GSD","expr":"if(D_GSD in ('Needed','Non-compliant'), 'open', if(D_GSD in ('Completed','Compliant'), 'closed', if(Q_GSD in ('Needed','Non-compliant'), 'open', if(Q_GSD in ('Completed','Compliant'), 'closed', blank))))","note":"Scorecard GSD overrides PtListQuality"},
{"rule":"derive","field":"S_CBP","expr":"if(Q_CBP in ('Needed','Non-compliant'), 'open', if(Q_CBP in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_OMW","expr":"if(Q_OMW in ('Needed','Non-compliant'), 'open', if(Q_OMW in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_KED","expr":"if(Diabetic and Egfr >= 100 and Uacr >= 100, 'closed', if(D_KED in ('Needed','Non-compliant'), 'open', if(D_KED in ('Completed','Compliant'), 'closed', if(Q_KED in ('Needed','Non-compliant'), 'open', if(Q_KED in ('Completed','Compliant'), 'closed', blank)))))","note":"Scorecard KED overrides PtListQuality; eGFR and uACR both at 100% close it"},
{"rule":"derive","field":"S_SPC","expr":"if(Q_SPC in ('Needed','Non-compliant'), 'open', if(Q_SPC in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_MAD","expr":"if(MedAdhDM in ('Needed','Non-compliant'), 'open', if(MedAdhDM in ('Completed','Compliant'), 'closed', if(Q_MAD in ('Needed','Non-compliant'), 'open', if(Q_MAD in ('Completed','Compliant'), 'closed', blank))))","note":"Scorecard Med Adherence DM overrides PtListQuality"},
{"rule":"derive","field":"S_MAC","expr":"if(Q_MAC in ('Needed','Non-compliant'), 'open', if(Q_MAC in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_MAH","expr":"if(Q_MAH in ('Needed','Non-compliant'), 'open', if(Q_MAH in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_SUPD","expr":"if(Q_SUPD in ('Needed','Non-compliant'), 'open', if(Q_SUPD in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_COB","expr":"if(Q_COB in ('Needed','Non-compliant'), 'open', if(Q_COB in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"S_POLY","expr":"if(Q_POLY in ('Needed','Non-compliant'), 'open', if(Q_POLY in ('Completed','Compliant'), 'closed', blank))"},
{"rule":"derive","field":"OpenHedis","expr":"join(', ', if(S_BCS = 'open', 'BCS', blank), if(S_COLO = 'open', 'COLO', blank), if(S_EED = 'open', 'EED', blank), if(S_GSD = 'open', 'GSD', blank), if(S_CBP = 'open', 'CBP', blank), if(S_OMW = 'open', 'OMW', blank), if(S_KED = 'open', 'KED', blank), if(S_SPC = 'open', 'SPC', blank), if(S_MAD = 'open', 'MAD', blank), if(S_MAC = 'open', 'MAC', blank), if(S_MAH = 'open', 'MAH', blank), if(S_SUPD = 'open', 'SUPD', blank), if(S_COB = 'open', 'COB', blank), if(S_POLY = 'open', 'POLY', blank))","note":"Compared as a set: the reports may list the overridden measures in another order"},
{"rule":"derive","field":"ClosedHedis","expr":"join(', ', if(S_BCS = 'closed', 'BCS', blank), if(S_COLO = 'closed', 'COLO', blank), if(S_EED = 'closed', 'EED', blank), if(S_GSD = 'closed', 'GSD', blank), if(S_CBP = 'closed', 'CBP', blank), if(S_OMW = 'closed', 'OMW', blank), if(S_KED = 'closed', 'KED', blank), if(S_SPC = 'closed', 'SPC', blank), if(S_MAD = 'closed', 'MAD', blank), if(S_MAC = 'closed', 'MAC', blank), if(S_MAH = 'closed', 'MAH', blank), if(S_SUPD = 'closed', 'SUPD', blank), if(S_COB = 'closed', 'COB', blank), if(S_POLY = 'closed', 'POLY', blank))"},
{"rule":"derive","field":"LastVisit","expr":"max(LastQemNonPcp, LastAcv, PcpVisit)"},
{"rule":"derive","field":"SerialLast","expr":"latestDate(PcpVisitDates, today())","note":"Latest PCP visit date in Serial Scheduling that is not in the future"},
{"rule":"derive","field":"A1c","expr":"number(replace(A1cText, '%$', ''))"},
{"rule":"derive","field":"EyeOpen","expr":"Diabetic and EyeStatus = 'Open'"},
{"rule":"derive","field":"EgfrNeeded","expr":"Diabetic and Egfr < 100"},
{"rule":"derive","field":"UacrNeeded","expr":"Diabetic and Uacr < 100"},
{"rule":"derive","field":"TotalGaps","expr":"coalesce(number(replace(TotalCareGaps, '%$', '')), 0)"},
{"rule":"derive","field":"IncompleteAttest","expr":"coalesce(number(replace(IncompletedAttest, '%$', '')), 0)"},
{"rule":"derive","field":"OpenIcdCount","expr":"countItems(OpenICD, ',', '0')"},
{"rule":"derive","field":"Tcm","expr":"upper(trim(TcmFlag)) in ('1','Y','YES','TRUE') or coalesce(number(replace(TcmFlag, '%$', '')), 0) > 0"},
{"rule":"filter","keep":"InExport or HighRisk","note":"The reports list a provider's Export patients plus high-risk patients not found in Export"},
{"rule":"sort","field":"Provider"},
{"rule":"sort","field":"LastName"},
{"rule":"sort","field":"FirstName"},
{"rule":"output","field":"NPI"},
{"rule":"output","field":"Provider"},
{"rule":"output","field":"Location"},
{"rule":"output","field":"RiskPool"},
{"rule":"output","field":"MemberID"},
{"rule":"output","field":"MRN"},
{"rule":"output","field":"FirstName","header":"First"},
{"rule":"output","field":"LastName","header":"Last"},
{"rule":"output","field":"BirthDate","header":"DOB"},
{"rule":"output","field":"ACOR"},
{"rule":"output","field":"ACVDue"},
{"rule":"output","field":"Attest"},
{"rule":"output","field":"OpenICD"},
{"rule":"output","field":"NewPatient"},
{"rule":"output","field":"OpenHedis"},
{"rule":"output","field":"ClosedHedis"},
{"rule":"output","field":"HighRisk"},
{"rule":"output","field":"LastVisit"},
{"rule":"output","field":"SerialLast"},
{"rule":"output","field":"A1cDate"},
{"rule":"output","field":"FutureVisits"},
{"rule":"output","field":"NextAppt"},
{"rule":"output","field":"NextApptSpecialty"},
{"rule":"output","field":"A1c"},
{"rule":"output","field":"A1cText"},
{"rule":"output","field":"Diabetic"},
{"rule":"output","field":"EyeOpen"},
{"rule":"output","field":"EyeDate"},
{"rule":"output","field":"EgfrNeeded"},
{"rule":"output","field":"UacrNeeded"},
{"rule":"output","field":"MedAdhDM"},
{"rule":"output","field":"OmwDue"},
{"rule":"output","field":"TotalGaps"},
{"rule":"output","field":"IncompleteAttest"},
{"rule":"output","field":"OpenIcdCount"},
{"rule":"output","field":"Tcm"},
{"rule":"output","field":"Payer"},
{"rule":"output","field":"Phone"},
{"rule":"output","field":"ActiveDate"},
{"rule":"output","field":"PcpVisit"},
{"rule":"output","field":"NextAcv"},
{"rule":"view","view":"By provider","groupBy":["Location","Provider"]},
{"rule":"measure","view":"By provider","name":"Patients"},
{"rule":"measure","view":"By provider","name":"ACV due","where":"ACVDue"},
{"rule":"measure","view":"By provider","name":"High risk","where":"HighRisk"},
{"rule":"measure","view":"By provider","name":"Diabetic","where":"Diabetic"},
{"rule":"measure","view":"By provider","name":"With open HEDIS","where":"OpenHedis is not blank"},
{"rule":"view","view":"By location","groupBy":["Location"]},
{"rule":"measure","view":"By location","name":"Patients"},
{"rule":"measure","view":"By location","name":"ACV due","where":"ACVDue"},
{"rule":"measure","view":"By location","name":"High risk","where":"HighRisk"},
{"rule":"measure","view":"By location","name":"Diabetic","where":"Diabetic"},
{"rule":"measure","view":"By location","name":"With open HEDIS","where":"OpenHedis is not blank"},
{"rule":"view","view":"By risk pool","groupBy":["RiskPool"]},
{"rule":"measure","view":"By risk pool","name":"Patients"},
{"rule":"measure","view":"By risk pool","name":"ACV due","where":"ACVDue"},
{"rule":"measure","view":"By risk pool","name":"High risk","where":"HighRisk"},
{"rule":"measure","view":"By risk pool","name":"Diabetic","where":"Diabetic"},
{"rule":"measure","view":"By risk pool","name":"With open HEDIS","where":"OpenHedis is not blank"}
]
'@
$script:CmbParityFields=[ordered]@{MemberID='text';MRN='text';First='text';Last='text';DOB='date';ACOR='bool';ACVDue='bool';Attest='bool';OpenICD='text';NewPatient='bool';OpenHedis='list';ClosedHedis='list';HighRisk='bool';LastVisit='date';SerialLast='date';A1cDate='date';FutureVisits='text';NextAppt='date';NextApptSpecialty='text';A1c='number';A1cText='text';Diabetic='bool';EyeOpen='bool';EyeDate='date';EgfrNeeded='bool';UacrNeeded='bool';MedAdhDM='text';OmwDue='date';TotalGaps='number';IncompleteAttest='number';OpenIcdCount='number';Tcm='bool';Payer='text';Phone='text';ActiveDate='date';PcpVisit='date';NextAcv='date'}
function Set-CmbProviderAnalysisConfig{
 # Writes the config that reproduces the reports' patient facts; an existing one is kept in its history first.
 $r=Read-CmbConfigText $script:CmbProviderAnalysisConfig;if(@($r.errors).Count){throw ('The built-in config could not be read: '+$r.errors[0].message)}
 $sv=Save-CmbConfig $script:CmbPaConfigName @($r.rules) 'written from the report code';return (Get-CmbConfigModel $sv.name)
}
function ConvertTo-CmbParityValue($V,[string]$Kind){
 # One value in the form both sides are compared in: blank as '', true/false, dates as yyyy-MM-dd, numbers without trailing zeros, lists sorted.
 if($null -eq $V){return ''};if($V -is [string] -and [string]::IsNullOrWhiteSpace($V)){return ''}
 switch($Kind){
  'bool'{if($V -is [bool]){if($V){return 'TRUE'};return 'FALSE'};$b=[CmbRt]::Bool($V);if($null -eq $b){return [CmbRt]::Text($V).Trim()};if($b){return 'TRUE'};return 'FALSE'}
  'date'{$d=[CmbRt]::Date($V);if($null -eq $d){return [CmbRt]::Text($V).Trim()};$t=[DateTime]$d;if($t.TimeOfDay.Ticks -eq 0){return $t.ToString('yyyy-MM-dd')};return $t.ToString('yyyy-MM-dd HH:mm')}
  'number'{$n=[CmbRt]::Num($V);if($null -eq $n){return [CmbRt]::Text($V).Trim()};return ([double]$n).ToString('0.######',[Globalization.CultureInfo]::InvariantCulture)}
  'list'{return ((@(([CmbRt]::Text($V)) -split ','|ForEach-Object{$_.Trim()}|Where-Object{$_})|Sort-Object) -join ',')}
 }
 return [CmbRt]::Text($V).Trim()
}
function New-CmbParityRun($Body){
 $name=[string](Get-P $Body 'name' '');$c=Read-CmbConfigFile $name;if(@($c.errors).Count){throw ('The config file could not be read (line '+$c.errors[0].line+': '+$c.errors[0].message+').')}
 if(@(Get-ProviderProfiles|Where-Object{$null -ne $_ -and [string](Get-P $_ 'npi' '') -match '^\d{10}$'}).Count -eq 0){throw 'There are no saved provider profiles; the reports (and so the comparison) are built per provider.'}
 foreach($r in @(Get-CmbRuns)){if([string](Get-P $r 'state' '') -in @('Starting','Running') -and !(Update-CmbLiveness $r)){throw 'A combine run is already running; wait for it to finish.'}}
 $sha=[Security.Cryptography.SHA256]::Create();try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($c.text))) -replace '-','')}finally{$sha.Dispose()}
 $id=[Guid]::NewGuid().ToString('N');$path=Get-CmbRunPath $id
 $run=[ordered]@{cmbRunVersion=1;runId=$id;kind='parity';state='Starting';percent=1;stage='Launching the side-by-side worker';requestedUtc=[DateTime]::UtcNow.ToString('o');configName=$c.name;configPath=$c.path;configText=$c.text;configSha256=$hash;assignments=@();workerPid=0}
 Save-JsonAtomic $path $run;Log 'COMBINE_PARITY_REQUESTED' 'OK' $c.name
 if($script:CmbInline){Invoke-CmbRun $id}
 else{try{$null=Start-CmbWorker $id}catch{$run.state='Failed';$run.stage='Worker launch failed';$run['errorSummary']=$_.Exception.Message;Save-JsonAtomic $path $run;Log 'COMBINE_FAILED' 'FAILED' $_.Exception.Message}}
 return (Get-CmbRunModel $id)
}
function Invoke-CmbParity($Run,[string]$Path){
 # Combine side: the config on the current Provider Analysis files. Reports side: every saved profile's patient facts, read and joined by
 # the report code. Rows pair up on NPI + Member ID; each field is compared after ConvertTo-CmbParityValue.
 Set-RunProgress $Run $Path 2 'Loading the current Provider Analysis files into Combine'
 $imp=@(Import-CmbCanonicalFiles);if($imp.Count -eq 0){throw 'No Provider Analysis source files have been imported yet (Sources tab).'}
 $own=@{};foreach($x in $imp){$own[[string]$x.token]=$true}
 $cfg=Read-CmbConfigText ([string]$Run.configText);$rules=@($cfg.rules);$files=@(Get-CmbLiveFiles);$plan0=Get-CmbPlan $rules
 $choices=@{};foreach($s in $plan0.sources){foreach($c in @(Find-CmbFileCandidates $s $files)){if($own.ContainsKey($c.token)){$choices[$s.key]=$c.token+'|'+$c.sheet;break}}}
 $assign=Get-CmbAssignments $plan0 $files $choices;$plan=Get-CmbPlan $rules (Get-CmbTablesFromAssignments $assign) -Strict
 if($plan.errorCount){$first=@($plan.problems|Where-Object{$_.severity -eq 'error'}|Select-Object -First 3|ForEach-Object{$(if($_.rule){'row '+$_.rule+': '}else{''})+$_.message});throw ('The config has problems to fix first: '+($first -join ' | '))}
 $picks=@{};$asg=@();foreach($k in @($assign.Keys)){$p=$assign[$k].pick;if(!$p){continue};$f=Get-CmbFile $p.token;$picks[$k]=@{pick=$p;file=$f};$asg+=[ordered]@{source=$assign[$k].source;token=$p.token;sheet=$p.sheet;file=$p.file;sha256=[string]$f.sha256;rows=$p.rows}}
 Set-P $Run 'assignments' $asg
 $sw=[Diagnostics.Stopwatch]::StartNew()
 $res=Invoke-CmbEngine $plan $picks {param([int]$pc,[string]$st)Set-RunProgress $Run $Path ([int](4+$pc*0.5)) ('Combine: '+$st)}
 $cmbSeconds=[Math]::Round($sw.Elapsed.TotalSeconds,1)
 $hdr=@{};foreach($c in $res.outCols){if(!$hdr.ContainsKey($c.header)){$hdr[$c.header]=$c.key}}
 foreach($need in @('NPI','MemberID')){if(!$hdr.ContainsKey($need)){throw ('The config''s output needs a column headed "'+$need+'" to pair rows with the reports.')}}
 $fields=@($script:CmbParityFields.Keys|Where-Object{$hdr.ContainsKey($_)});$skipped=@($script:CmbParityFields.Keys|Where-Object{!$hdr.ContainsKey($_)})
 # Reports side
 $sw=[Diagnostics.Stopwatch]::StartNew()
 $profiles=@(Get-ProviderProfiles|Where-Object{$null -ne $_ -and [string](Get-P $_ 'npi' '') -match '^\d{10}$'})
 $read=Read-GroupSources $profiles $Run $Path 56 24
 Set-RunProgress $Run $Path 82 'Reports: joining each provider''s patients'
 $orig=[ordered]@{};$origDup=0;$names=@{}
 foreach($p in $profiles){
  $npi=[string]$p.npi;$names[$npi]=[string](Get-P $p 'displayName' $npi)
  $f=Get-ProviderFactsFromGroups $read.groups $p
  foreach($pt in @($f.facts.patients)){$k=$npi+'|'+[CmbRt]::Key($pt.MemberID);if($orig.Contains($k)){$origDup++}else{$orig[$k]=$pt}}
 }
 $repSeconds=[Math]::Round($sw.Elapsed.TotalSeconds,1)
 # Combine rows
 Set-RunProgress $Run $Path 90 'Comparing'
 $FV=$res.F;$npiA=$FV[$hdr['NPI']];$midA=$FV[$hdr['MemberID']];$cmb=@{};$cmbDup=0
 foreach($q in $res.order){$k=[CmbRt]::Text($npiA[$q]).Trim()+'|'+[CmbRt]::Key($midA[$q]);if($cmb.ContainsKey($k)){$cmbDup++}else{$cmb[$k]=$q}}
 $diffCount=[ordered]@{};foreach($f in $fields){$diffCount[$f]=0};$diffs=[Collections.Generic.List[object]]::new();$matched=0;$identical=0;$onlyRep=[Collections.Generic.List[object]]::new()
 foreach($k in @($orig.Keys)){
  $pt=$orig[$k];$npi=$k.Substring(0,$k.IndexOf('|'))
  if(!$cmb.ContainsKey($k)){$onlyRep.Add([ordered]@{npi=$npi;provider=[string]$names[$npi];memberId=[string]$pt.MemberID;name=(([string]$pt.Last).Trim()+', '+([string]$pt.First).Trim()).Trim(', ')});continue}
  $q=$cmb[$k];$matched++;$same=$true
  foreach($f in $fields){
   $kind=$script:CmbParityFields[$f];$a=ConvertTo-CmbParityValue $pt.$f $kind;$b=ConvertTo-CmbParityValue $FV[$hdr[$f]][$q] $kind
   if($a -cne $b){$same=$false;$diffCount[$f]++;if($diffs.Count -lt 100000){$diffs.Add([ordered]@{npi=$npi;provider=[string]$names[$npi];memberId=[string]$pt.MemberID;field=$f;reports=$a;combine=$b})}}
  }
  if($same){$identical++}
 }
 $onlyCmb=[Collections.Generic.List[object]]::new()
 foreach($k in @($cmb.Keys)){if(!$orig.Contains($k)){$q=$cmb[$k];$npi=$k.Substring(0,$k.IndexOf('|'));$nm='';if($hdr.ContainsKey('Last')){$nm=[CmbRt]::Text($FV[$hdr['Last']][$q])};if($hdr.ContainsKey('First')){$nm=($nm+', '+[CmbRt]::Text($FV[$hdr['First']][$q])).Trim(', ')};$onlyCmb.Add([ordered]@{npi=$npi;provider=[string]$names[$npi];memberId=[CmbRt]::Text($midA[$q]);name=$nm})}}
 # Workbooks: the usual Combine output, and the side-by-side comparison.
 Set-RunProgress $Run $Path 94 'Writing the workbooks'
 $stamp=(Get-Date).ToString('yyyy-MM-dd HHmmss');$base=(Get-SafeFileName ([string]$Run.configName))+' - '+$stamp
 $out=Join-Path $script:Paths.CombineOutputs ($base+'.xlsx');$n=2;while(Test-Path -LiteralPath $out){$out=Join-Path $script:Paths.CombineOutputs ($base+' ('+$n+').xlsx');$n++}
 Write-CmbWorkbook $res $out $Run
 $pout=Join-Path $script:Paths.CombineOutputs ('Side by side - '+$base+'.xlsx');$n=2;while(Test-Path -LiteralPath $pout){$pout=Join-Path $script:Paths.CombineOutputs ('Side by side - '+$base+' ('+$n+').xlsx');$n++}
 $sum=[ordered]@{providers=$profiles.Count;reportRows=$orig.Count;combineRows=$cmb.Count;matched=$matched;identical=$identical;onlyReports=$onlyRep.Count;onlyCombine=$onlyCmb.Count;reportDuplicates=$origDup;combineDuplicates=$cmbDup;differentRows=($matched-$identical);fieldsCompared=$fields.Count;fieldsNotCompared=$skipped;combineSeconds=$cmbSeconds;reportSeconds=$repSeconds;differences=@(foreach($f in $fields){[ordered]@{field=$f;rows=$diffCount[$f]}});samples=@($diffs|Select-Object -First 200);onlyReportsSample=@($onlyRep|Select-Object -First 100);onlyCombineSample=@($onlyCmb|Select-Object -First 100)}
 $pkg=New-Object OfficeOpenXml.ExcelPackage;$used=@{}
 try{
  $items=@(@('Config',[string]$Run.configName),@('Built',(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')),@('Providers with profiles',[double]$sum.providers),@('Patients in the reports',[double]$sum.reportRows),@('Patients in Combine',[double]$sum.combineRows),@('Paired (same NPI and Member ID)',[double]$sum.matched),@('Paired and identical in every compared field',[double]$sum.identical),@('Paired with differences',[double]$sum.differentRows),@('Only in the reports',[double]$sum.onlyReports),@('Only in Combine',[double]$sum.onlyCombine),@('Repeated in the reports (same provider and Member ID)',[double]$sum.reportDuplicates),@('Repeated in Combine',[double]$sum.combineDuplicates),@('Fields compared',($fields -join ', ')),@('Fields not compared (no output column)',($skipped -join ', ')))
  $null=Add-CmbSheet $pkg $used 'Summary' @('Item','Value') @(foreach($x in $items){,([object[]]$x)}) $null @(52,90)
  $null=Add-CmbSheet $pkg $used 'Field differences' @('Field','Rows that differ') @(foreach($f in $fields){,([object[]]@($f,[double]$diffCount[$f]))})
  $null=Add-CmbSheet $pkg $used 'Differences' @('Provider','NPI','Member ID','Field','Reports','Combine') @(foreach($d in $diffs){,([object[]]@($d.provider,$d.npi,$d.memberId,$d.field,$d.reports,$d.combine))})
  $null=Add-CmbSheet $pkg $used 'Only in the reports' @('Provider','NPI','Member ID','Name') @(foreach($d in $onlyRep){,([object[]]@($d.provider,$d.npi,$d.memberId,$d.name))})
  $null=Add-CmbSheet $pkg $used 'Only in Combine' @('Provider','NPI','Member ID','Name') @(foreach($d in $onlyCmb){,([object[]]@($d.provider,$d.npi,$d.memberId,$d.name))})
  [IO.File]::WriteAllBytes($pout,$pkg.GetAsByteArray())
 }finally{$pkg.Dispose()}
 return @{result=(Get-CmbRunSummary $res);parity=$sum;outputPath=$out;parityPath=$pout}
}
# --- Combine web: page models, request handling and the three pages (Files, Rules, Output).
function Get-CmbLiveFiles{return @(Get-CmbFiles|Where-Object{Test-Path -LiteralPath ([string](Get-P $_ 'path' ''))})}
function Get-CmbFilesModel{
 $files=@(Get-CmbLiveFiles);$uses=@{}
 foreach($c in @(Get-CmbConfigs)){try{$cf=Read-CmbConfigFile $c.name;$plan=Get-CmbPlan $cf.rules;foreach($s in $plan.sources){foreach($cand in @(Find-CmbFileCandidates $s $files)){$k=$cand.token;if(!$uses.ContainsKey($k)){$uses[$k]=@()};$uses[$k]+=($c.name+' / '+$s.id)}}}catch{}}
 $list=@(foreach($f in $files){$m=Get-CmbFileModel ([string]$f.token);$m['usedBy']=@($(if($uses.ContainsKey([string]$f.token)){$uses[[string]$f.token]}else{@()}));$m})
 return [ordered]@{files=$list;folder=$script:Paths.CombineFiles;keep=$script:CmbFileKeep;configs=@(Get-CmbConfigs|ForEach-Object{$_.name})}
}
function Remove-CmbFileRequest([string]$Token){
 $f=Json (Get-CmbFilePath $Token);if(!$f){throw 'That file is not loaded.'}
 foreach($r in @(Get-CmbRuns)){if([string](Get-P $r 'state' '') -in @('Starting','Running')){foreach($a in @(Get-P $r 'assignments' @())){if([string](Get-P $a 'token' '') -eq $Token){throw 'A run is reading this file; wait for it to finish.'}}}}
 Remove-CmbFileRecord $f;Log 'COMBINE_FILE_REMOVED' 'OK' ([string]$f.originalName);return [ordered]@{removed=$true}
}
function Get-CmbChoices($Body){$c=@{};$ch=Get-P $Body 'choices' $null;if($ch){foreach($p in $ch.PSObject.Properties){$c[$p.Name.ToLowerInvariant()]=[string]$p.Value}};return $c}
function Get-CmbConfigModel([string]$Name){
 $c=Read-CmbConfigFile $Name;$check=Get-CmbCheck $c.rules @(Get-CmbLiveFiles)
 return [ordered]@{name=$c.name;path=$c.path;readErrors=@($c.errors);check=$check;history=@(Get-CmbConfigHistory $c.name)}
}
function Get-CmbCheckModel($Body){
 # Checks rules sent by the page (the editor's current text) without saving them.
 $r=Read-CmbConfigText ([string](Get-P $Body 'text' ''))
 if(@($r.errors).Count -and @($r.rules).Count -eq 0){return [ordered]@{readErrors=@($r.errors);check=$null}}
 return [ordered]@{readErrors=@($r.errors);check=(Get-CmbCheck @($r.rules) @(Get-CmbLiveFiles) (Get-CmbChoices $Body))}
}
function Save-CmbConfigRequest($Body){
 $name=Get-CmbConfigName ([string](Get-P $Body 'name' ''));$from=[string](Get-P $Body 'renameFrom' '');$create=[bool](Get-P $Body 'create' $false)
 $r=Read-CmbConfigText ([string](Get-P $Body 'text' ''));if(@($r.errors).Count){$e=$r.errors[0];throw ('Not saved: '+$(if($e.line){'line '+$e.line+': '}else{''})+$e.message)}
 $exists=Test-Path -LiteralPath (Get-CmbConfigPath $name)
 if(($create -or ($from -and $from -ne $name)) -and $exists){throw ('A config named "'+$name+'" already exists; choose another name.')}
 $sv=Save-CmbConfig $name @($r.rules) $(if($from -and $from -ne $name){'renamed from '+$from}elseif($create){'created'}else{'saved'})
 if($from -and $from -ne $name -and (Test-Path -LiteralPath (Get-CmbConfigPath $from))){Remove-CmbConfig $from}
 return (Get-CmbConfigModel $sv.name)
}
function Get-CmbAnswerModel($Body){
 # An LLM answer pasted back: read row by row, compared with the editor's rules, and checked; nothing is saved here.
 $ans=Read-CmbConfigText ([string](Get-P $Body 'answer' ''));$cur=Read-CmbConfigText ([string](Get-P $Body 'text' ''))
 if(@($ans.rules).Count -eq 0){return [ordered]@{readErrors=@($ans.errors);rules=@();diff=@();check=$null}}
 $d=Compare-CmbRules @($cur.rules) @($ans.rules)
 return [ordered]@{readErrors=@($ans.errors);rules=@($ans.rules|ForEach-Object{ConvertTo-CmbRuleText $_});diff=$d;check=(Get-CmbCheck @($ans.rules) @(Get-CmbLiveFiles))}
}
function Get-CmbDiffModel($Body){$a=Read-CmbConfigText ([string](Get-P $Body 'before' ''));$b=Read-CmbConfigText ([string](Get-P $Body 'after' ''));$d=Compare-CmbRules @($a.rules) @($b.rules);return [ordered]@{diff=$d}}
function Get-CmbGuideModel{
 return [ordered]@{
  guide=@($script:CmbGuide);expressions=@($script:CmbExprGuide);answer=@($script:CmbAnswerGuide)
  ruleTypes=@(foreach($t in @($script:CmbRuleSpec.Keys)){$s=$script:CmbRuleSpec[$t];[ordered]@{rule=$t;purpose=$s.purpose;props=@(foreach($k in @($s.props.Keys)){[ordered]@{name=$k;text=$s.props[$k]}});example=$s.example}})
  how=@(foreach($k in @($script:CmbHowText.Keys)){[ordered]@{name=$k;text=$script:CmbHowText[$k]}});clean=@(foreach($k in @($script:CmbCleanText.Keys)){[ordered]@{name=$k;text=$script:CmbCleanText[$k]}});functions=@(foreach($k in @($script:CmbFunctionText.Keys)){[ordered]@{name=$k;text=$script:CmbFunctionText[$k]}})
 }
}
function Get-CmbSetupModel([string]$Name,[hashtable]$Choices){
 $c=Read-CmbConfigFile $Name;$check=Get-CmbCheck $c.rules @(Get-CmbLiveFiles) $Choices
 $problems=@(foreach($r in $check.rules){foreach($p in $r.problems){if($p.severity -eq 'error'){[ordered]@{row=$r.index;message=$p.message}}}})+@(foreach($p in $check.general){if($p.severity -eq 'error'){[ordered]@{row=0;message=$p.message}}})
 return [ordered]@{name=$c.name;readErrors=@($c.errors);sources=$check.sources;errors=$check.errors;warnings=$check.warnings;problems=$problems;views=$check.views;fields=$check.fields}
}
function Get-CmbOutputPageModel{
 $runs=@(Get-CmbRuns);$latest=$null;$latestParity=$null
 foreach($r in $runs){$k=[string](Get-P $r 'kind' '');if($k -eq 'parity'){if(!$latestParity){$latestParity=Get-CmbRunModel ([string]$r.runId)}}elseif(!$latest){$latest=Get-CmbRunModel ([string]$r.runId)}}
 $canon=@(foreach($src in @(Registry)){if(Test-Path -LiteralPath (Join-Path $script:Paths.CanonicalCurrent ([string]$src.canonicalFileName))){[string]$src.sourceKey}})
 return [ordered]@{configs=@(Get-CmbConfigs|ForEach-Object{$_.name});latest=$latest;latestParity=$latestParity;paConfig=$script:CmbPaConfigName;canonicalSources=$canon;profiles=@(Get-OverviewProviders).Count;history=@(foreach($r in $runs){if([string](Get-P $r 'state' '') -in @('Completed','Failed')){Get-CmbRunModel ([string]$r.runId) -Summary}});folder=$script:Paths.CombineOutputs}
}
function Send-CmbJson($Context,$Value){Send $Context 200 'application/json' ($Value|ConvertTo-Json -Depth 25)}
function Invoke-CmbRequest($Context,[string]$Method,[string]$Path){
 # Every Combine route; returns $false for paths it does not own.
 $q=$Context.Request.QueryString
 if($Method -eq 'GET'){
  switch($Path){
   '/combine'{Send $Context 200 'text/html; charset=utf-8' (Get-CmbPage 'files');return $true}
   '/combine/rules'{Send $Context 200 'text/html; charset=utf-8' (Get-CmbPage 'rules');return $true}
   '/combine/output'{Send $Context 200 'text/html; charset=utf-8' (Get-CmbPage 'output');return $true}
   '/combine/output.xlsx'{Send-CmbOutput $Context ([string]$q['id']);return $true}
   '/combine/parity.xlsx'{Send-CmbOutput $Context ([string]$q['id']) 'parityPath';return $true}
   '/api/combine/files'{Send-CmbJson $Context (Get-CmbFilesModel);return $true}
   '/api/combine/file'{Send-CmbJson $Context (Get-CmbFileModel ([string]$q['token']) -Full);return $true}
   '/api/combine/configs'{Send-CmbJson $Context ([ordered]@{configs=@(Get-CmbConfigs);folder=$script:Paths.CombineConfigs});return $true}
   '/api/combine/config'{Send-CmbJson $Context (Get-CmbConfigModel ([string]$q['name']));return $true}
   '/api/combine/version'{Send-CmbJson $Context ([ordered]@{text=(Read-CmbConfigVersion ([string]$q['name']) ([string]$q['v']))});return $true}
   '/api/combine/guide'{Send-CmbJson $Context (Get-CmbGuideModel);return $true}
   '/api/combine/setup'{$ch=@{};$raw=[string]$q['choices'];if($raw){$o=$raw|ConvertFrom-Json;foreach($p in $o.PSObject.Properties){$ch[$p.Name.ToLowerInvariant()]=[string]$p.Value}};Send-CmbJson $Context (Get-CmbSetupModel ([string]$q['name']) $ch);return $true}
   '/api/combine/output'{Send-CmbJson $Context (Get-CmbOutputPageModel);return $true}
   '/api/combine/run'{Send-CmbJson $Context (Get-CmbRunModel ([string]$q['id']));return $true}
  }
  return $false
 }
 if($Method -eq 'POST'){
  switch($Path){
   '/api/combine/upload'{Send-CmbJson $Context (Receive-CmbUpload $Context);return $true}
   '/api/combine/file-remove'{Send-CmbJson $Context (Remove-CmbFileRequest ([string]$q['token']));return $true}
   '/api/combine/check'{Send-CmbJson $Context (Get-CmbCheckModel (Read-BodyJson $Context));return $true}
   '/api/combine/save'{Send-CmbJson $Context (Save-CmbConfigRequest (Read-BodyJson $Context));return $true}
   '/api/combine/delete'{Remove-CmbConfig ([string]$q['name']);Send-CmbJson $Context ([ordered]@{deleted=$true});return $true}
   '/api/combine/answer'{Send-CmbJson $Context (Get-CmbAnswerModel (Read-BodyJson $Context));return $true}
   '/api/combine/diff'{Send-CmbJson $Context (Get-CmbDiffModel (Read-BodyJson $Context));return $true}
   '/api/combine/prompt'{$b=Read-BodyJson $Context;Send-CmbJson $Context (Get-CmbPrompt ([string](Get-P $b 'name' '')) ([string](Get-P $b 'text' '')) ([string](Get-P $b 'request' '')) ([string[]]@(Get-P $b 'tokens' @())) ([bool](Get-P $b 'examples' $false)));return $true}
   '/api/combine/draft'{$b=Read-BodyJson $Context;$n=Get-CmbConfigName ([string](Get-P $b 'name' ''));$rules=New-CmbDraft $n ([string[]]@(Get-P $b 'tokens' @()));Send-CmbJson $Context ([ordered]@{name=$n;text=(Format-CmbConfigText $rules)});return $true}
   '/api/combine/run'{Send-CmbJson $Context (New-CmbRun (Read-BodyJson $Context));return $true}
   '/api/combine/parity'{Send-CmbJson $Context (New-CmbParityRun (Read-BodyJson $Context));return $true}
   '/api/combine/pa-config'{Send-CmbJson $Context (Set-CmbProviderAnalysisConfig);return $true}
   '/api/combine/import-canonical'{Send-CmbJson $Context ([ordered]@{imported=@(Import-CmbCanonicalFiles)});return $true}
  }
 }
 return $false
}
$script:CmbCss=@'
body{font:14px Segoe UI,Arial;margin:0;background:#f4f7fb;color:#172033}header{background:#17365d;color:white;padding:22px 22px 12px}header h1{margin:0 0 4px}main{padding:16px 22px;max-width:1600px}
.sub{display:flex;gap:6px;margin:0 0 6px}.sub a{padding:7px 14px;border-radius:16px;text-decoration:none;color:#17365d;background:#e4e9f0;font-weight:600}.sub a.on{background:#17365d;color:#fff}.sub .cfg{margin-left:auto;align-self:center;color:#667085}
.card{background:white;border:1px solid #dce4ef;border-radius:8px;padding:16px 18px;margin:14px 0}.card h2{margin:0 0 8px;font-size:17px;color:#17365d}.muted{color:#667085}.error{color:#a61b1b}.warn{color:#9a5b00}.ok{color:#1d7a3a}
.row{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:8px 0}select,input[type=text],textarea{font:inherit;padding:5px 7px;border:1px solid #c8d2e0;border-radius:5px}textarea{width:100%;box-sizing:border-box;font:12.5px Consolas,monospace}
button,a.btn{padding:6px 11px;cursor:pointer;background:#1769aa;color:#fff;border:0;border-radius:5px;font:inherit;text-decoration:none;display:inline-block}button:disabled{opacity:.5;cursor:default}button.alt,a.btn.alt{background:#e4e9f0;color:#172033}button.go{background:#1d7a3a;font-weight:600;padding:9px 16px}button.del{background:#a61b1b}button.mini{padding:2px 7px;font-size:12px}
table.grid{border-collapse:collapse;width:100%}table.grid th,table.grid td{padding:6px 8px;border-bottom:1px solid #e6ecf3;text-align:left;vertical-align:top}table.grid th{background:#eaf1f8}.scroll table.grid th{position:sticky;top:0}table.grid td.n,table.grid th.n{text-align:right}
.scroll{overflow:auto;max-height:560px;border:1px solid #e6ecf3;border-radius:6px}.wide td{white-space:nowrap}
tr.err td{background:#fdecec}tr.wrn td{background:#fff6e5}tr.sel td{background:#eef6ff}
.badge{display:inline-block;padding:1px 7px;border-radius:9px;font-size:12px;font-weight:600;background:#e4e9f0;color:#17365d}.b-source{background:#dbeafe}.b-map{background:#e0f2f1}.b-filter{background:#fde2e2}.b-key{background:#fef3c7}.b-field{background:#ede9fe}.b-derive{background:#fce7f3}.b-lookup{background:#e0e7ff}.b-sort,.b-output{background:#ecfccb}.b-view,.b-measure{background:#ffedd5}.b-config{background:#17365d;color:#fff}
.prob{font-size:12.5px;margin:2px 0}.json{font:12px Consolas,monospace;white-space:pre-wrap;word-break:break-all;color:#334155}
.tiles{display:grid;grid-template-columns:repeat(6,minmax(130px,1fr));gap:10px;margin:10px 0}.tile{border:1px solid #dce4ef;border-radius:8px;padding:9px 12px}.tile b{display:block;font-size:21px;color:#1f6d7a}.tile small{color:#667085}
.modal{position:fixed;inset:0;background:#0007;display:none;align-items:center;justify-content:center;z-index:9}.modal.on{display:flex}.box{background:white;border-radius:8px;padding:18px;width:min(1100px,94vw);max-height:90vh;overflow:auto}
.drop{border:2px dashed #9fb3cc;border-radius:8px;padding:18px;text-align:center;color:#667085}.drop.over{background:#eef6ff;border-color:#1769aa}
.chips{display:flex;gap:5px;flex-wrap:wrap}.chip{border:1px solid #c8d2e0;border-radius:14px;padding:2px 9px;cursor:pointer;background:#f8fafc;font-size:12.5px}.chip.on{background:#17365d;color:#fff;border-color:#17365d}
.diff td{font:12px Consolas,monospace;white-space:pre-wrap;word-break:break-all}.d-added{color:#1d7a3a}.d-removed{color:#a61b1b}.d-changed,.d-reordered{color:#9a5b00}
.steps{counter-reset:s;margin:6px 0 0;padding-left:20px}.steps li{margin:3px 0}
code{background:#eef2f7;padding:1px 5px;border-radius:3px;word-break:break-all}details summary{cursor:pointer;color:#17365d;font-weight:600}
@media(max-width:900px){.tiles{grid-template-columns:1fr 1fr}}
'@
$script:CmbJs=@'
const el=id=>document.getElementById(id);
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
async function api(u,o){const r=await fetch(u,o);const t=await r.text();let j=null;try{j=t?JSON.parse(t):null}catch(e){throw Error('Server returned invalid JSON: '+t.slice(0,200))}if(!r.ok)throw Error((j&&j.error)||('Request failed ('+r.status+')'));return j}
function post(u,b){return api(u,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b||{})})}
function when(iso){if(!iso)return '';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleString()}
function num(n){return n==null?'':Number(n).toLocaleString()}
function plural(n,w){return num(n)+' '+w+(n===1?'':'s')}
function arr(x){return x==null?[]:(Array.isArray(x)?x:[x])}
function curConfig(){let n='';try{n=new URL(location).searchParams.get('config')||localStorage.getItem('cmb.config')||''}catch(e){}return n}
function rememberConfig(n){try{localStorage.setItem('cmb.config',n||'')}catch(e){}document.querySelectorAll('.sub a[data-p]').forEach(a=>{a.href=a.dataset.p+(n?'?config='+encodeURIComponent(n):'')});const c=el('subCfg');if(c)c.textContent=n?'Config: '+n:''}
function modal(html){el('modalBody').innerHTML=html;el('modal').classList.add('on')}
function closeModal(){el('modal').classList.remove('on')}
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal()});
function diffHtml(d){d=arr(d);if(!d.length)return '<p class="ok">No differences.</p>';return '<table class="grid diff"><thead><tr><th>Change</th><th>Rule</th><th>Before</th><th>After</th></tr></thead><tbody>'+d.map(x=>'<tr><td class="d-'+esc(x.change)+'"><b>'+esc(x.change)+'</b></td><td>'+esc(x.identity)+'</td><td>'+esc(x.before)+'</td><td>'+esc(x.after)+'</td></tr>').join('')+'</tbody></table>'}
'@
function Get-CmbPage([string]$Which){
 $sub='<nav class="sub"><a data-p="/combine" href="/combine"'+$(if($Which -eq 'files'){' class="on"'}else{''})+'>1. Files</a><a data-p="/combine/rules" href="/combine/rules"'+$(if($Which -eq 'rules'){' class="on"'}else{''})+'>2. Rules</a><a data-p="/combine/output" href="/combine/output"'+$(if($Which -eq 'output'){' class="on"'}else{''})+'>3. Output</a><span class="cfg" id="subCfg"></span></nav>'
 $title=@{files='Combine - Files';rules='Combine - Rules';output='Combine - Output'}[$Which]
 $body=@{files=$script:CmbFilesHtml;rules=$script:CmbRulesHtml;output=$script:CmbOutputHtml}[$Which]
 $html='<!doctype html><html><head><meta charset="utf-8"><title>'+$title+'</title><style>'+$script:TabCss+$script:RunPageCss+$script:CmbCss+'</style><script>'+$script:CmbJs+'</script></head><body><header><h1>Combine</h1><p>Draft 6.0 - load exports, write the rules that combine them into one row per patient, and build the output with a full audit trail</p></header>'+(Get-NavHtml 'combine')+'<main>'+$sub+$body+'</main><div id="modal" class="modal" onclick="if(event.target===this)closeModal()"><div class="box"><div id="modalBody"></div><p><button class="alt" onclick="closeModal()">Close</button></p></div></div></body></html>'
 return $html.Replace('__FOLDER__',[Net.WebUtility]::HtmlEncode($script:Paths.Combine)).Replace('__FILEKEEP__',[string]$script:CmbFileKeep)
}
$script:CmbFilesHtml=@'
<section class="card"><h2>Load exports</h2>
<p class="muted">Excel (.xlsx) or text (.csv, .txt, .tsv) database exports. Each file is fingerprinted when it loads: the header row is found (even under title rows), and every column gets its type, fill rate, distinct values, value shapes and common values. The newest __FILEKEEP__ files are kept in <code>__FOLDER__</code>; loading the same file again reuses it.</p>
<div id="drop" class="drop">Drop files here, or <button id="pick">Choose files...</button><input id="file" type="file" accept=".xlsx,.csv,.txt,.tsv" multiple hidden></div>
<div id="queue"></div>
<div class="row"><button class="alt" id="canon">Load the current Provider Analysis source files</button><span class="muted">Adds the seven files imported on the Sources tab, under their usual names, so a config can use them.</span></div>
</section>
<section class="card"><h2>Loaded files</h2>
<div class="row"><button id="draft" disabled>Draft a config from the selected files...</button><span class="muted">Writes a first set of rules (sources, column maps, the patient key, field priorities, output columns and views) for you or your LLM to refine on the Rules page.</span></div>
<div id="files"><span class="muted">Loading...</span></div>
<p class="muted">Used by lists the config sources whose match columns this file has. Example values stay on this computer; they are sent to the LLM only if you tick that option on the Rules page.</p>
</section>
<script>
let files=[];const open={};
function sheetText(f){return arr(f.sheets).map(s=>(f.extension==='.xlsx'?esc(s.name)+': ':'')+plural(s.rows,'row')+', '+plural(s.columnCount,'column')+(s.headerRow>1?' (headers on row '+s.headerRow+')':'')).join('<br>')}
function render(){
 if(!files.length){el('files').innerHTML='<p class="muted">No files loaded yet.</p>';syncDraft();return}
 el('files').innerHTML='<table class="grid"><thead><tr><th></th><th>File</th><th>Tables</th><th>Layout id</th><th>File date</th><th>Loaded</th><th>Used by</th><th></th></tr></thead><tbody>'+files.map(f=>'<tr><td><input type="checkbox" class="pickf" value="'+esc(f.token)+'"></td><td><b>'+esc(f.originalName)+'</b><br><span class="muted">'+num(Math.round(f.byteLength/1024))+' KB'+(f.seconds!=null?', read in '+f.seconds+' s':'')+'</span></td><td>'+sheetText(f)+'</td><td>'+arr(f.sheets).map(s=>'<code>'+esc(s.layoutId)+'</code>').join('<br>')+'</td><td>'+esc(when(f.fileDateUtc))+'</td><td>'+esc(when(f.loadedUtc))+'</td><td>'+(arr(f.usedBy).length?arr(f.usedBy).map(esc).join('<br>'):'<span class="muted">-</span>')+'</td><td style="white-space:nowrap"><button class="alt mini" data-cols="'+esc(f.token)+'">'+(open[f.token]?'Hide columns':'Columns')+'</button> <button class="alt mini" data-rm="'+esc(f.token)+'">Remove</button></td></tr>'+(open[f.token]?'<tr><td></td><td colspan="7" id="cols-'+esc(f.token)+'">Loading...</td></tr>':'')).join('')+'</tbody></table>';
 for(const t in open){if(open[t])showColumns(t)}
 syncDraft();
}
function syncDraft(){el('draft').disabled=!document.querySelectorAll('.pickf:checked').length}
let showEx=false;
async function showColumns(t){
 const c=el('cols-'+t);if(!c)return;
 try{const f=await api('/api/combine/file?token='+encodeURIComponent(t));
  c.innerHTML='<label class="muted"><input type="checkbox" id="ex-'+esc(t)+'"'+(showEx?' checked':'')+'> show example values</label>'+arr(f.sheets).map(s=>'<h3 style="margin:10px 0 4px;font-size:14px">'+esc(s.name)+' <span class="muted">- headers on row '+s.headerRow+', '+plural(s.rows,'data row')+(s.delimiter?', '+esc(s.delimiter)+'-separated, '+esc(s.encoding):'')+(s.malformedLines?', <span class="warn">'+plural(s.malformedLines,'malformed line')+'</span>':'')+'</span></h3><div class="scroll"><table class="grid"><thead><tr><th>Col</th><th>Header</th><th>Type</th><th class="n">Filled</th><th class="n">Blank</th><th class="n">Distinct</th><th>Shapes</th><th>Common values</th><th>Range</th>'+(showEx?'<th>Examples</th>':'')+'</tr></thead><tbody>'+arr(s.columns).map(x=>'<tr><td>'+esc(x.letter)+'</td><td><b>'+esc(x.header)+'</b>'+(x.unique?' <span class="badge b-key" title="every filled value is different">unique</span>':'')+'</td><td>'+esc(x.type)+'</td><td class="n">'+num(x.filled)+'</td><td class="n">'+num(x.blank)+'</td><td class="n">'+num(x.distinct)+'</td><td>'+arr(x.patterns).map(p=>'<code>'+esc(p.shape)+'</code> '+Math.round(p.share*100)+'%').join('<br>')+'</td><td>'+arr(x.common).slice(0,8).map(v=>esc(v.value)+' <span class="muted">('+num(v.rows)+')</span>').join('<br>')+(arr(x.common).length>8?'<br><span class="muted">+'+(arr(x.common).length-8)+' more</span>':'')+'</td><td>'+(x.min?esc(x.min)+' to '+esc(x.max):'')+'</td>'+(showEx?'<td>'+arr(x.examples).map(esc).join('<br>')+'</td>':'')+'</tr>').join('')+'</tbody></table></div>').join('');
  const cb=el('ex-'+t);if(cb)cb.onchange=()=>{showEx=cb.checked;for(const k in open){if(open[k])showColumns(k)}};
 }catch(e){c.innerHTML='<span class="error">'+esc(e.message)+'</span>'}
}
el('files').addEventListener('click',async e=>{
 const b=e.target.closest('button');if(!b)return;
 if(b.dataset.cols){open[b.dataset.cols]=!open[b.dataset.cols];render();return}
 if(b.dataset.rm){const f=files.find(x=>x.token===b.dataset.rm);if(!confirm('Remove '+(f?f.originalName:'this file')+' from the loaded files?'))return;try{await post('/api/combine/file-remove?token='+encodeURIComponent(b.dataset.rm));delete open[b.dataset.rm];await load()}catch(err){alert(err.message)}}
});
el('files').addEventListener('change',e=>{if(e.target.classList.contains('pickf'))syncDraft()});
async function load(){const m=await api('/api/combine/files');files=arr(m.files);render()}
async function uploadAll(list){
 for(const f of list){
  const id='q'+Math.random().toString(36).slice(2);el('queue').insertAdjacentHTML('beforeend','<div class="row" id="'+id+'"><b>'+esc(f.name)+'</b> <span class="muted">loading and fingerprinting...</span></div>');
  try{const j=await api('/api/combine/upload',{method:'POST',headers:{'X-File-Name':encodeURIComponent(f.name),'X-File-Timestamp':String(f.lastModified)},body:f});
   el(id).innerHTML='<b>'+esc(f.name)+'</b> <span class="ok">loaded</span> <span class="muted">'+arr(j.sheets).map(s=>plural(s.rows,'row')+' x '+plural(s.columnCount,'column')).join('; ')+'</span>';open[j.token]=list.length===1}
  catch(e){el(id).innerHTML='<b>'+esc(f.name)+'</b> <span class="error">'+esc(e.message)+'</span>'}
  await load();
 }
}
el('pick').onclick=()=>{el('file').value='';el('file').click()};
el('file').onchange=()=>uploadAll([...el('file').files]);
el('canon').onclick=async()=>{el('canon').disabled=true;el('queue').insertAdjacentHTML('beforeend','<div class="row" id="canonNote"><span class="muted">Loading the Provider Analysis source files...</span></div>');try{const j=await post('/api/combine/import-canonical');el('canonNote').innerHTML=arr(j.imported).length?'<span class="ok">Loaded</span> '+arr(j.imported).map(x=>esc(x.file)).join(', '):'<span class="warn">No source files have been imported on the Sources tab yet.</span>';await load()}catch(e){el('canonNote').innerHTML='<span class="error">'+esc(e.message)+'</span>'}el('canon').disabled=false;el('canonNote').removeAttribute('id')};
const drop=el('drop');['dragenter','dragover'].forEach(n=>drop.addEventListener(n,e=>{e.preventDefault();drop.classList.add('over')}));['dragleave','drop'].forEach(n=>drop.addEventListener(n,e=>{e.preventDefault();drop.classList.remove('over')}));drop.addEventListener('drop',e=>uploadAll([...e.dataTransfer.files]));
el('draft').onclick=async()=>{
 const tokens=[...document.querySelectorAll('.pickf:checked')].map(x=>x.value);if(!tokens.length)return;
 const name=prompt('Name for the new config:','Combined patients');if(!name)return;
 try{const j=await post('/api/combine/draft',{name:name,tokens:tokens});sessionStorage.setItem('cmb.draft',JSON.stringify(j));location.href='/combine/rules?draft=1'}catch(e){alert(e.message)}
};
rememberConfig(curConfig());
load().catch(e=>{el('files').innerHTML='<p class="error">'+esc(e.message)+'</p>'});
</script>
'@
$script:CmbRulesHtml=@'
<section class="card"><h2>Config</h2>
<div class="row"><select id="cfg"></select><button class="alt" id="new">New</button><button class="alt" id="dup">Duplicate</button><button class="alt" id="ren">Rename</button><button class="alt" id="dl">Download .json</button><button class="del" id="del">Delete</button><span id="state" class="muted"></span></div>
<div class="row"><button class="go" id="save" disabled>Save</button><button class="alt" id="revert" disabled>Undo unsaved changes</button><span id="path" class="muted"></span></div>
<div id="general"></div>
</section>
<section class="card"><h2>Rules</h2>
<p class="muted">One row per rule, applied in the order explained in the reference below. Edit opens the rule's JSON with the reference for its type; changes are checked at once against the loaded files and are kept only when you Save. Every save keeps the previous version.</p>
<div class="row"><div class="chips" id="chips"></div><input type="text" id="find" placeholder="Find..." style="width:180px"><select id="addType"><option value="">Add a rule...</option></select></div>
<div class="scroll" style="max-height:640px"><table class="grid" id="rulesTable"><thead><tr><th>#</th><th>Rule</th><th>What it does</th><th>Note</th><th>Check</th><th></th></tr></thead><tbody id="rules"><tr><td colspan="6" class="muted">Loading...</td></tr></tbody></table></div>
</section>
<section class="card"><h2>Source files</h2><div id="sources" class="muted">Loading...</div><p class="muted">Each source uses the newest loaded file that has its match columns. Load files on the <a href="/combine">Files</a> page; choose between several matches on the <a href="/combine/output">Output</a> page.</p></section>
<section class="card"><h2>Work with your LLM</h2>
<ol class="steps">
<li>Describe the change (or leave it blank for a review), choose the files to describe, and copy the request. It is one JSON document laid out one item per line: it explains how this program works, every rule type and function, the current config with its problems, and each file's columns.</li>
<li>Paste it into cell A1 of a blank Excel sheet (or straight into the LLM chat) and ask the LLM to follow it. It answers with a new sheet: column A, one rule per row.</li>
<li>Copy column A of that sheet and paste it below, then Check the answer. You see exactly which rules it added, changed or removed, and any problems, before anything is saved.</li>
</ol>
<div class="row" style="align-items:flex-start"><textarea id="ask" rows="3" placeholder="What should change? e.g. Add a view by risk pool and region with AWV due counts. Leave blank to have the LLM review the config."></textarea></div>
<div class="row"><b>Describe files:</b><span id="askFiles" class="muted">Loading...</span></div>
<div class="row"><label><input type="checkbox" id="askEx"> include example values from each column <span class="warn">(these are patient data; only for an LLM approved for PHI)</span></label></div>
<div class="row"><button id="copy">Copy request to clipboard</button><button class="alt" id="dlAsk">Download request (.txt)</button><span id="askNote" class="muted"></span></div>
<textarea id="answer" rows="6" placeholder="Paste the LLM's answer here (column A copied from its sheet, or the JSON it wrote)."></textarea>
<div class="row"><button id="checkAns">Check answer</button><span id="ansNote" class="muted"></span></div>
<div id="ansOut"></div>
</section>
<section class="card"><h2>Versions</h2><div id="history" class="muted">Loading...</div></section>
<section class="card"><details><summary>Reference: how rules work</summary><div id="guide" class="muted">Loading...</div></details></section>
<script>
let configs=[],cfgName='',rules=[],savedText='',check=null,guide=null,files=[],typeFilter='',history=[],answerRules=null,lastCheckText='';
const ORDER=['config','source','map','filter','key','field','derive','lookup','sort','output','view','measure'];
function textOf(list){return '[\r\n'+list.map(r=>JSON.stringify(r)).join(',\r\n')+'\r\n]\r\n'}
function dirty(){return textOf(rules)!==savedText}
function sync(){
 const d=dirty();el('save').disabled=!d;el('revert').disabled=!d||!savedText;el('state').innerHTML=d?'<span class="warn"><b>Unsaved changes</b></span>':(cfgName?'<span class="ok">Saved</span>':'');
 ['dup','ren','dl','del'].forEach(i=>el(i).disabled=!cfgName&&i!=='dl');
}
window.addEventListener('beforeunload',e=>{if(dirty()){e.preventDefault();e.returnValue=''}});
function ruleType(r){return String((r&&r.rule)||'').toLowerCase()}
function probHtml(ps){return arr(ps).map(p=>'<div class="prob '+(p.severity==='error'?'error':'warn')+'">'+esc(p.message)+'</div>').join('')}
function render(){
 const rows=check?arr(check.rules):[];const q=el('find').value.trim().toLowerCase();
 const chips=['']. concat(ORDER.filter(t=>rules.some(r=>ruleType(r)===t)));
 el('chips').innerHTML=chips.map(t=>'<span class="chip'+(t===typeFilter?' on':'')+'" data-t="'+t+'">'+(t?esc(t)+' ('+rules.filter(r=>ruleType(r)===t).length+')':'All ('+rules.length+')')+'</span>').join('');
 el('rules').innerHTML=rules.map((r,i)=>{
  const c=rows[i]||{},t=ruleType(r);if(typeFilter&&t!==typeFilter)return '';
  const sum=c.summary||'',note=(r&&r.note)||'';if(q&&(sum+' '+note+' '+JSON.stringify(r)).toLowerCase().indexOf(q)<0)return '';
  const ps=arr(c.problems),bad=ps.some(p=>p.severity==='error'),wn=!bad&&ps.length;
  return '<tr class="'+(bad?'err':wn?'wrn':'')+'"><td class="n">'+(i+1)+'</td><td><span class="badge b-'+esc(t)+'">'+esc(t||'?')+'</span></td><td>'+esc(sum)+'</td><td>'+esc(note)+'</td><td>'+(ps.length?probHtml(ps):'<span class="ok">ok</span>')+'</td><td style="white-space:nowrap"><button class="mini" data-edit="'+i+'">Edit</button> <button class="alt mini" data-up="'+i+'" title="Move up">&uarr;</button><button class="alt mini" data-down="'+i+'" title="Move down">&darr;</button> <button class="alt mini" data-rm="'+i+'" title="Delete">&times;</button></td></tr>'}).join('')||'<tr><td colspan="6" class="muted">No rules'+(typeFilter||q?' match':'')+'.</td></tr>';
 const g=check?arr(check.general):[];
 el('general').innerHTML=(check?'<p>'+(check.errors?'<span class="error"><b>'+plural(check.errors,'problem')+' to fix before a run</b></span>':'<span class="ok"><b>Ready to run</b></span>')+(check.warnings?' &middot; <span class="warn">'+plural(check.warnings,'warning')+'</span>':'')+' &middot; '+plural(rules.length,'rule')+(arr(check.views).length?' &middot; views: '+arr(check.views).map(esc).join(', '):'')+'</p>':'')+probHtml(g);
 el('sources').innerHTML=check&&arr(check.sources).length?'<table class="grid"><thead><tr><th>Source</th><th>Role</th><th>Match columns</th><th>File it would use</th><th class="n">Rows</th><th>Other matches</th></tr></thead><tbody>'+arr(check.sources).map(s=>'<tr><td><b>'+esc(s.source)+'</b><br><span class="muted">'+esc(s.label)+'</span></td><td>'+esc(s.role)+'</td><td>'+arr(s.match).map(esc).join(', ')+'</td><td>'+(s.file?esc(s.file)+(s.sheet&&/\.xlsx$/i.test(s.file)?' <span class="muted">sheet '+esc(s.sheet)+'</span>':''):'<span class="'+(s.required?'error':'warn')+'">no loaded file has these columns</span>')+'</td><td class="n">'+(s.file?num(s.rows):'')+'</td><td>'+(arr(s.candidates).length>1?arr(s.candidates).slice(1).map(x=>esc(x.file)).join('<br>'):'')+'</td></tr>').join('')+'</tbody></table>':'<span class="muted">No source rules yet.</span>';
 sync();
}
let timer=null;
function changed(){render();clearTimeout(timer);timer=setTimeout(recheck,250)}
async function recheck(){const t=textOf(rules);lastCheckText=t;try{const j=await post('/api/combine/check',{text:t});if(t!==lastCheckText)return;check=j.check;render()}catch(e){el('general').innerHTML='<p class="error">'+esc(e.message)+'</p>'}}
function loadRules(list,saved,name){rules=list;savedText=saved;cfgName=name;check=null;rememberConfig(name);el('path').innerHTML='';fillConfigs();render();recheck()}
async function openConfig(name){
 try{const m=await api('/api/combine/config?name='+encodeURIComponent(name));
  const list=arr(m.check.rules).map(r=>r.json?JSON.parse(r.json):{});
  rules=list;cfgName=m.name;savedText=textOf(list);check=m.check;history=arr(m.history);rememberConfig(cfgName);
  el('path').innerHTML='<code>'+esc(m.path)+'</code>'+(arr(m.readErrors).length?' <span class="error">'+arr(m.readErrors).map(e=>'line '+e.line+': '+esc(e.message)).join('; ')+' (those lines are not shown and would be dropped on save)</span>':'');
  fillConfigs();render();renderHistory();renderAskFiles()}
 catch(e){el('general').innerHTML='<p class="error">'+esc(e.message)+'</p>'}
}
function fillConfigs(){const s=el('cfg');s.innerHTML=(cfgName&&!configs.includes(cfgName)?'<option value="'+esc(cfgName)+'">'+esc(cfgName)+' (not saved yet)</option>':'')+(configs.length||cfgName?'':'<option value="">No configs yet</option>')+configs.map(n=>'<option value="'+esc(n)+'">'+esc(n)+'</option>').join('');s.value=cfgName}
async function loadConfigs(){const m=await api('/api/combine/configs');configs=arr(m.configs).map(c=>c.name)}
el('cfg').onchange=async()=>{const n=el('cfg').value;if(dirty()&&!confirm('Discard unsaved changes to '+(cfgName||'this config')+'?')){el('cfg').value=cfgName;return}if(n)await openConfig(n)};
el('chips').onclick=e=>{const c=e.target.closest('.chip');if(!c)return;typeFilter=c.dataset.t;render()};
el('find').oninput=()=>render();
el('rules').addEventListener('click',e=>{
 const b=e.target.closest('button');if(!b)return;
 if(b.dataset.edit!=null){editRule(+b.dataset.edit);return}
 if(b.dataset.up!=null){const i=+b.dataset.up;if(i>0){[rules[i-1],rules[i]]=[rules[i],rules[i-1]];changed()}return}
 if(b.dataset.down!=null){const i=+b.dataset.down;if(i<rules.length-1){[rules[i+1],rules[i]]=[rules[i],rules[i+1]];changed()}return}
 if(b.dataset.rm!=null){const i=+b.dataset.rm;if(confirm('Delete rule '+(i+1)+'?\n'+JSON.stringify(rules[i]))){rules.splice(i,1);changed()}}
});
function specHtml(t){
 if(!guide)return '';const s=arr(guide.ruleTypes).find(x=>x.rule===t);if(!s)return '<p class="error">Unknown rule type "'+esc(t)+'". Rule types: '+ORDER.join(', ')+'.</p>';
 return '<p><b>'+esc(t)+'</b>: '+esc(s.purpose)+'</p><table class="grid"><tbody>'+arr(s.props).map(p=>'<tr><td><code>'+esc(p.name)+'</code></td><td>'+esc(p.text)+'</td></tr>').join('')+'<tr><td><code>note</code></td><td>optional. Why the rule exists; shown in the audit.</td></tr></tbody></table><p class="muted">Example: <code>'+esc(s.example)+'</code></p>'+(t==='field'?'<p class="muted">how: '+arr(guide.how).map(h=>'<b>'+esc(h.name)+'</b> '+esc(h.text)).join('; ')+'</p>':'')+(t==='map'?'<p class="muted">clean steps: '+arr(guide.clean).map(h=>'<code>'+esc(h.name)+'</code> '+esc(h.text)).join('; ')+'</p>':'');
}
function editRule(i,isNew){
 const r=rules[i];
 modal('<h2 style="margin-top:0">'+(isNew?'New':'Edit')+' rule '+(i+1)+'</h2><textarea id="ruleText" rows="12"></textarea><div class="row"><span id="ruleNote" class="muted"></span></div><div class="row"><button id="ruleApply">Apply</button><button class="alt" id="ruleCancel">Cancel</button></div><div id="ruleSpec"></div>');
 const ta=el('ruleText');ta.value=JSON.stringify(r,null,2);
 const upd=()=>{try{const o=JSON.parse(ta.value);if(!o||typeof o!=='object'||Array.isArray(o))throw Error('a rule is one JSON object { ... }');el('ruleNote').className='ok';el('ruleNote').textContent='Valid JSON';el('ruleApply').disabled=false;el('ruleSpec').innerHTML=specHtml(ruleType(o))}catch(e){el('ruleNote').className='error';el('ruleNote').textContent=e.message;el('ruleApply').disabled=true}};
 ta.oninput=upd;upd();ta.focus();
 el('ruleApply').onclick=()=>{rules[i]=JSON.parse(ta.value);closeModal();changed()};
 el('ruleCancel').onclick=()=>{if(isNew){rules.splice(i,1);changed()}closeModal()};
}
el('addType').onchange=()=>{
 const t=el('addType').value;el('addType').value='';if(!t||!guide)return;
 const s=arr(guide.ruleTypes).find(x=>x.rule===t);let o={rule:t};try{o=JSON.parse(s.example)}catch(e){}
 let at=rules.length;for(let i=rules.length-1;i>=0;i--){if(ORDER.indexOf(ruleType(rules[i]))<=ORDER.indexOf(t)){at=i+1;break}if(i===0)at=0}
 rules.splice(at,0,o);typeFilter='';render();editRule(at,true);
};
el('save').onclick=async()=>{
 let name=cfgName;if(!name){name=prompt('Name for this config:','');if(!name)return}
 try{const m=await post('/api/combine/save',{name:name,text:textOf(rules),create:!configs.includes(name)});await loadConfigs();await openConfig(m.name)}catch(e){alert(e.message)}
};
el('revert').onclick=()=>{if(!confirm('Undo every unsaved change?'))return;openConfig(cfgName)};
el('new').onclick=()=>{if(dirty()&&!confirm('Discard unsaved changes?'))return;const n=prompt('Name for the new config:','');if(!n)return;if(configs.includes(n)){alert('A config named '+n+' already exists.');return}history=[];renderHistory();loadRules([{rule:'config',name:n,note:'New config'}],'',n)};
el('dup').onclick=async()=>{const n=prompt('Name for the copy:',cfgName+' copy');if(!n)return;try{const m=await post('/api/combine/save',{name:n,text:textOf(rules),create:true});await loadConfigs();await openConfig(m.name)}catch(e){alert(e.message)}};
el('ren').onclick=async()=>{const n=prompt('New name:',cfgName);if(!n||n===cfgName)return;try{const m=await post('/api/combine/save',{name:n,text:textOf(rules),renameFrom:cfgName});await loadConfigs();await openConfig(m.name)}catch(e){alert(e.message)}};
el('del').onclick=async()=>{if(!confirm('Delete config '+cfgName+'? Its file moves to the history folder.'))return;try{await post('/api/combine/delete?name='+encodeURIComponent(cfgName));savedText='';rules=[];await loadConfigs();if(configs.length)await openConfig(configs[0]);else loadRules([],'','')}catch(e){alert(e.message)}};
function download(name,text,type){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:type}));a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},500)}
el('dl').onclick=()=>download((cfgName||'config')+'.json',textOf(rules),'application/json');
function renderAskFiles(){
 const used=new Set(check?arr(check.sources).map(s=>s.file).filter(x=>x):[]);
 el('askFiles').innerHTML=files.length?files.map(f=>'<label style="margin-right:12px"><input type="checkbox" class="askf" value="'+esc(f.token)+'"'+(used.size===0||used.has(f.originalName)?' checked':'')+'> '+esc(f.originalName)+'</label>').join(''):'no files loaded (load them on the Files page so the LLM can see their columns)';
}
async function buildAsk(){const tokens=[...document.querySelectorAll('.askf:checked')].map(x=>x.value);return await post('/api/combine/prompt',{name:cfgName,text:textOf(rules),request:el('ask').value,tokens:tokens,examples:el('askEx').checked})}
async function copyText(t){try{await navigator.clipboard.writeText(t);return true}catch(e){const ta=document.createElement('textarea');ta.value=t;ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.select();let ok=false;try{ok=document.execCommand('copy')}catch(x){}ta.remove();return ok}}
el('copy').onclick=async()=>{el('askNote').textContent='Building...';try{const j=await buildAsk();const ok=await copyText(j.text);el('askNote').innerHTML=(ok?'<span class="ok">Copied</span> ':'<span class="error">Could not copy; use Download</span> ')+num(j.lines)+' lines, '+num(j.characters)+' characters ('+plural(j.rules,'rule')+', '+plural(j.files,'file')+'). Paste into cell A1 of a blank sheet.'}catch(e){el('askNote').innerHTML='<span class="error">'+esc(e.message)+'</span>'}};
el('dlAsk').onclick=async()=>{try{const j=await buildAsk();download('Combine request - '+(cfgName||'config')+'.txt',j.text,'text/plain')}catch(e){alert(e.message)}};
el('checkAns').onclick=async()=>{
 const t=el('answer').value;if(!t.trim()){el('ansNote').textContent='Paste the answer first.';return}
 el('ansNote').textContent='Checking...';el('ansOut').innerHTML='';answerRules=null;
 try{const j=await post('/api/combine/answer',{text:textOf(rules),answer:t});
  const re=arr(j.readErrors);
  if(!arr(j.rules).length){el('ansNote').innerHTML='<span class="error">No rules could be read.</span>';el('ansOut').innerHTML=re.map(e=>'<div class="prob error">'+(e.line?'Row '+e.line+': ':'')+esc(e.message)+'</div>').join('');return}
  answerRules=arr(j.rules).map(x=>JSON.parse(x));const c=j.check||{};
  el('ansNote').innerHTML=plural(answerRules.length,'rule')+' read'+(re.length?', <span class="error">'+plural(re.length,'row')+' could not be read</span>':'')+' &middot; '+(c.errors?'<span class="error">'+plural(c.errors,'problem')+'</span>':'<span class="ok">no problems</span>')+(c.warnings?', <span class="warn">'+plural(c.warnings,'warning')+'</span>':'');
  el('ansOut').innerHTML=re.map(e=>'<div class="prob error">Row '+e.line+': '+esc(e.message)+'</div>').join('')+'<h3 style="font-size:14px">Changes compared with the editor</h3>'+diffHtml(j.diff)+(c.errors||c.warnings?'<h3 style="font-size:14px">Problems in the answer</h3>'+arr(c.rules).filter(r=>arr(r.problems).length).map(r=>'<div><b>Rule '+r.index+'</b> '+esc(r.summary)+probHtml(r.problems)+'</div>').join('')+probHtml(c.general):'')+'<div class="row"><button class="go" id="useAns">Use this answer in the editor</button><span class="muted">Nothing is saved until you select Save.</span></div>';
  el('useAns').onclick=()=>{rules=answerRules;changed();el('ansOut').innerHTML='<p class="ok">The answer is in the editor. Review the rules above, then Save.</p>';window.scrollTo({top:0,behavior:'smooth'})};
 }catch(e){el('ansNote').innerHTML='<span class="error">'+esc(e.message)+'</span>'}
};
function renderHistory(){
 el('history').innerHTML=!cfgName?'<span class="muted">Not saved yet.</span>':history.length?'<table class="grid"><thead><tr><th>Saved before</th><th class="n">Rules</th><th></th></tr></thead><tbody>'+history.map(h=>'<tr><td>'+esc(when(h.savedUtc))+(/deleted/.test(h.version)?' <span class="warn">(deleted)</span>':'')+'</td><td class="n">'+num(h.rules)+'</td><td><a href="#" data-v="'+esc(h.version)+'" data-a="view">View</a> &middot; <a href="#" data-v="'+esc(h.version)+'" data-a="cmp">Compare with the editor</a> &middot; <a href="#" data-v="'+esc(h.version)+'" data-a="restore">Restore into the editor</a></td></tr>').join('')+'</tbody></table>':'<span class="muted">No earlier versions; each Save keeps the one before it.</span>';
}
el('history').addEventListener('click',async e=>{
 const a=e.target.closest('a[data-v]');if(!a)return;e.preventDefault();
 try{const j=await api('/api/combine/version?name='+encodeURIComponent(cfgName)+'&v='+encodeURIComponent(a.dataset.v));
  if(a.dataset.a==='view'){modal('<h2 style="margin-top:0">Version '+esc(a.dataset.v)+'</h2><pre class="json">'+esc(j.text)+'</pre>');return}
  if(a.dataset.a==='cmp'){const d=await post('/api/combine/diff',{before:j.text,after:textOf(rules)});modal('<h2 style="margin-top:0">From version '+esc(a.dataset.v)+' to the editor</h2>'+diffHtml(d.diff));return}
  if(!confirm('Load that version into the editor? Nothing is saved until you select Save.'))return;
  const c=await post('/api/combine/check',{text:j.text});if(!c.check){alert('That version could not be read.');return}rules=arr(c.check.rules).map(r=>JSON.parse(r.json));changed();
 }catch(err){alert(err.message)}
});
function guideHtml(g){
 return '<h3>How it works</h3>'+arr(g.guide).map(x=>'<p>'+esc(x)+'</p>').join('')+'<h3>Rule types</h3>'+arr(g.ruleTypes).map(s=>specHtml(s.rule)).join('<hr>')+'<h3>Expressions</h3>'+arr(g.expressions).map(x=>'<p>'+esc(x)+'</p>').join('')+'<table class="grid"><tbody>'+arr(g.functions).map(f=>'<tr><td><code>'+esc(f.name)+'</code></td><td>'+esc(f.text)+'</td></tr>').join('')+'</tbody></table><h3>Cleaning steps</h3><table class="grid"><tbody>'+arr(g.clean).map(f=>'<tr><td><code>'+esc(f.name)+'</code></td><td>'+esc(f.text)+'</td></tr>').join('')+'</tbody></table>';
}
(async()=>{
 try{
  guide=await api('/api/combine/guide');el('guide').className='';el('guide').innerHTML=guideHtml(guide);
  el('addType').innerHTML='<option value="">Add a rule...</option>'+ORDER.map(t=>'<option value="'+t+'">'+t+'</option>').join('');
  const fm=await api('/api/combine/files');files=arr(fm.files);
  await loadConfigs();
  let draft=null;if(new URL(location).searchParams.get('draft')){try{draft=JSON.parse(sessionStorage.getItem('cmb.draft')||'null')}catch(e){};sessionStorage.removeItem('cmb.draft')}
  if(draft){const c=await post('/api/combine/check',{text:draft.text});history=[];renderHistory();loadRules(arr(c.check&&c.check.rules).map(r=>JSON.parse(r.json)),'',draft.name);if(configs.includes(draft.name))el('state').innerHTML='<span class="warn">A config named '+esc(draft.name)+' exists; Save would replace it</span>';renderAskFiles();return}
  const want=curConfig();if(configs.length){await openConfig(configs.includes(want)?want:configs[0])}else{fillConfigs();render();renderHistory();renderAskFiles();el('general').innerHTML='<p class="muted">No configs yet. Load your exports on the Files page and use Draft a config, or select New.</p>'}
 }catch(e){el('general').innerHTML='<p class="error">'+esc(e.message)+'</p>'}
})();
</script>
'@
$script:CmbOutputHtml=@'
<section class="card"><h2>Run a config</h2>
<div class="row"><b>Config</b><select id="cfg"></select><a class="btn alt" id="editLink" href="/combine/rules">Edit rules</a></div>
<div id="setup" class="muted">Loading...</div>
<div class="row"><button class="go" id="run" disabled>Combine</button><span id="runNote" class="muted"></span></div>
</section>
<section class="card"><h2>Side by side with the provider reports</h2>
<p class="muted">Checks a config against the code that builds the provider reports. It loads the current Provider Analysis source files (Sources tab) into Combine, runs the config on them, builds every saved provider's patient facts exactly as the reports do, and pairs the two by NPI and Member ID, comparing every field. The <b>Provider Analysis facts</b> config is that report logic written as rules, so a clean comparison shows the rules reproduce the reports.</p>
<div class="row"><button class="alt" id="paCfg">Create or refresh the "Provider Analysis facts" config</button><button id="parity">Compare the selected config with the reports</button><span id="parityNote" class="muted"></span></div>
<div id="parityOut"></div></section>
<section class="card"><h2>Result</h2><div id="result"><span class="muted">Loading...</span></div></section>
<section class="card"><h2>Previous runs</h2><p class="muted">Every run, newest first. Each workbook is saved in <code id="folder"></code> with the patient sheet, one sheet per view, and audit sheets: every rule's effect, the files used (with their SHA-256), conflicts, key conflicts, excluded patients, the source of every value, and the exact config.</p><div id="history"><span class="muted">Loading...</span></div></section>
<script>
let configs=[],setup=null,current=null,history=[],timer=null,choices={},parityCur=null,paName='Provider Analysis facts',canon=[],profileCount=0;
function running(o){return !!o&&(o.state==='Starting'||o.state==='Running')}
function busy(){return running(current)||running(parityCur)}
function sync(){el('run').disabled=busy()||!setup||setup.errors>0||!!arr(setup.readErrors).length;el('run').textContent=running(current)?'Combining...':'Combine';el('parity').disabled=busy()||!el('cfg').value||!canon.length||!profileCount;el('paCfg').disabled=busy();el('parity').textContent=running(parityCur)?'Comparing...':'Compare the selected config with the reports'}
function showParity(o){
 parityCur=o;const box=el('parityOut');
 el('parityNote').textContent=(canon.length?canon.length+' Provider Analysis source files, ':'No Provider Analysis source files imported yet (Sources tab). ')+(profileCount?plural(profileCount,'provider profile'):'no provider profiles yet (Provider Wizard)');
 if(!o){box.innerHTML='';sync();return}
 if(running(o)){box.innerHTML='<p>Comparing '+esc(o.configName)+'... '+o.percent+'%</p><div class="progress"><span style="width:'+Math.max(0,Math.min(100,o.percent))+'%"></span></div><span class="muted">'+esc(o.stage)+'</span>';sync();return}
 if(o.state==='Failed'){box.innerHTML='<p class="error">The comparison failed: '+esc(o.errorSummary)+'</p>';sync();return}
 const p=o.parity||{};const clean=!p.differentRows&&!p.onlyReports&&!p.onlyCombine;
 let h='<div class="row"><b>'+esc(o.configName)+'</b><span class="muted">compared '+esc(when(o.completedUtc))+(o.seconds!=null?' in '+o.seconds+' s':'')+'</span>'+(o.parityUrl?'<a class="btn" href="'+esc(o.parityUrl)+'">Download the comparison</a>':'')+(o.outputUrl?'<a class="btn alt" href="'+esc(o.outputUrl)+'">Combine workbook</a>':'')+'</div>';
 h+='<p class="'+(clean?'ok':'warn')+'"><b>'+(clean?'Identical: every patient of every provider matches in all '+p.fieldsCompared+' fields.':'Differences found; see below and the Differences sheet.')+'</b></p>';
 h+='<div class="tiles">'+tile('Providers',num(p.providers))+tile('Patients in the reports',num(p.reportRows))+tile('Patients in Combine',num(p.combineRows))+tile('Identical',num(p.identical),'of '+num(p.matched)+' paired')+tile('Only in the reports',num(p.onlyReports))+tile('Only in Combine',num(p.onlyCombine))+'</div>';
 const d=arr(p.differences).filter(x=>x.rows>0);
 h+='<p class="muted">'+num(p.fieldsCompared)+' fields compared'+(arr(p.fieldsNotCompared).length?'; not compared (no output column): '+arr(p.fieldsNotCompared).map(esc).join(', '):'')+'. Reports took '+p.reportSeconds+' s, Combine '+p.combineSeconds+' s.</p>';
 if(d.length)h+='<h3>Fields that differ</h3>'+table(['Field','Patients'],d.map(x=>[x.field,num(x.rows)]),null,1)+'<h3>Differences <span class="muted">(first '+arr(p.samples).length+')</span></h3>'+table(['Provider','Member ID','Field','Reports','Combine'],arr(p.samples).map(x=>[x.provider,x.memberId,x.field,x.reports,x.combine]));
 if(arr(p.onlyReportsSample).length)h+='<h3>Only in the reports</h3>'+table(['Provider','Member ID','Name'],arr(p.onlyReportsSample).map(x=>[x.provider,x.memberId,x.name]));
 if(arr(p.onlyCombineSample).length)h+='<h3>Only in Combine</h3>'+table(['Provider','Member ID','Name'],arr(p.onlyCombineSample).map(x=>[x.provider,x.memberId,x.name]));
 box.innerHTML=h;sync();
}
el('paCfg').onclick=async()=>{if(configs.includes(paName)&&!confirm('Replace the "'+paName+'" config with the rules written from the report code? Its current version is kept in its history.'))return;try{await post('/api/combine/pa-config');if(!configs.includes(paName)){configs.push(paName);configs.sort()}el('cfg').innerHTML=configs.map(n=>'<option value="'+esc(n)+'">'+esc(n)+'</option>').join('');el('cfg').value=paName;choices={};await loadSetup();el('parityNote').textContent='"'+paName+'" written; select Compare to check it against the reports.'}catch(e){alert(e.message)}};
el('parity').onclick=async()=>{if(el('parity').disabled)return;el('parity').disabled=true;try{const o=await post('/api/combine/parity',{name:el('cfg').value});showParity(o);if(running(o))startPoll(o.runId)}catch(e){el('parityNote').innerHTML='<span class="error">'+esc(e.message)+'</span>';sync()}};
async function loadSetup(){
 const n=el('cfg').value;rememberConfig(n);el('editLink').href='/combine/rules?config='+encodeURIComponent(n);if(!n){setup=null;el('setup').innerHTML='<p class="muted">No configs yet; make one on the <a href="/combine/rules">Rules</a> page.</p>';sync();return}
 try{setup=await api('/api/combine/setup?name='+encodeURIComponent(n)+'&choices='+encodeURIComponent(JSON.stringify(choices)));renderSetup()}catch(e){setup=null;el('setup').innerHTML='<p class="error">'+esc(e.message)+'</p>'}sync();
}
function renderSetup(){
 const s=setup;let h='<table class="grid"><thead><tr><th>Source</th><th>Role</th><th>File</th><th class="n">Rows</th></tr></thead><tbody>'+arr(s.sources).map(x=>'<tr><td><b>'+esc(x.source)+'</b><br><span class="muted">'+esc(x.label)+'</span></td><td>'+esc(x.role)+(x.required?'':' <span class="muted">(optional)</span>')+'</td><td>'+(arr(x.candidates).length?'<select data-src="'+esc(x.source)+'">'+arr(x.candidates).map(c=>'<option value="'+esc(c.value)+'"'+(c.value===x.chosen?' selected':'')+'>'+esc(c.file)+(/\.xlsx$/i.test(c.file)?' - '+esc(c.sheet):'')+' (loaded '+esc(when(c.loadedUtc))+')</option>').join('')+'</select>':'<span class="'+(x.required?'error':'warn')+'">No loaded file has the columns '+arr(x.match).map(esc).join(', ')+'.</span> <a href="/combine">Load it</a>')+'</td><td class="n">'+(x.chosen?num(x.rows):'')+'</td></tr>').join('')+'</tbody></table>';
 if(arr(s.readErrors).length)h+='<p class="error">The config file has lines that cannot be read; fix them on the Rules page.</p>';
 if(s.errors)h+='<p class="error"><b>'+plural(s.errors,'problem')+' to fix first</b> (on the <a href="/combine/rules?config='+encodeURIComponent(s.name)+'">Rules</a> page):</p>'+arr(s.problems).slice(0,8).map(p=>'<div class="prob error">'+(p.row?'Rule '+p.row+': ':'')+esc(p.message)+'</div>').join('');
 else h+='<p class="ok">Ready: '+plural(arr(s.fields).length,'field')+(arr(s.views).length?', views '+arr(s.views).map(esc).join(', '):'')+(s.warnings?' <span class="warn">('+plural(s.warnings,'warning')+'; see Rules)</span>':'')+'</p>';
 el('setup').innerHTML=h;
}
el('setup').addEventListener('change',e=>{const s=e.target.closest('select[data-src]');if(!s)return;choices[s.dataset.src.toLowerCase()]=s.value;loadSetup()});
el('cfg').onchange=()=>{choices={};loadSetup()};
function tile(l,v,n){return '<div class="tile">'+esc(l)+'<b>'+v+'</b>'+(n?'<small>'+esc(n)+'</small>':'')+'</div>'}
function table(headers,rows,total,groups){return '<div class="scroll"><table class="grid wide"><thead><tr>'+arr(headers).map((h,i)=>'<th'+(groups!=null&&i>=groups?' class="n"':'')+'>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+arr(rows).map(r=>'<tr>'+arr(r).map((c,i)=>'<td'+(groups!=null&&i>=groups?' class="n"':'')+'>'+esc(c)+'</td>').join('')+'</tr>').join('')+(total?'<tr>'+arr(total).map((c,i)=>'<td'+(groups!=null&&i>=groups?' class="n"':'')+'><b>'+esc(c)+'</b></td>').join('')+'</tr>':'')+'</tbody></table></div>'}
function showResult(o){
 current=o;
 if(!o){el('result').innerHTML='<span class="muted">No runs yet.</span>';sync();return}
 if(busy()){el('result').innerHTML='<p>Combining '+esc(o.configName)+'... '+o.percent+'%</p><div class="progress"><span style="width:'+Math.max(0,Math.min(100,o.percent))+'%"></span></div><span class="muted">'+esc(o.stage)+'</span>';sync();return}
 if(o.state==='Failed'){el('result').innerHTML='<p class="error">The run failed: '+esc(o.errorSummary)+'</p><span class="muted">'+esc(o.configName)+' &middot; requested '+esc(when(o.requestedUtc))+'</span>';sync();return}
 const r=o.result||{};
 let h='<div class="row"><b>'+esc(o.configName)+'</b><span class="muted">built '+esc(when(o.completedUtc))+(o.seconds!=null?' in '+o.seconds+' s':'')+' &middot; today() = '+esc(r.asOf)+'</span>'+(o.outputUrl?'<a class="btn" href="'+esc(o.outputUrl)+'">Download workbook</a>':'<span class="warn">'+esc(o.outputError)+'</span>')+'</div>';
 h+='<div class="tiles">'+tile('Patients',num(r.patients),r.matched!==r.patients?num(r.matched)+' before patient filters':'')+tile('Rows used',num(r.records),'from '+plural(arr(r.sources).filter(s=>s.role==='patients').length,'source'))+tile('Rows without a key',num(r.noKey),'left out')+tile('Conflicts',num(r.conflicts),'patients where sources disagree')+tile('Key conflicts',num(r.keyConflicts),'joins refused')+tile('Excluded',num(r.excluded),'by patient filters')+'</div>';
 h+=(o.outputFile?'<p class="muted">Saved as <code>'+esc(o.outputFile)+'</code></p>':'');
 h+='<h3>'+esc(r.sheet||'Patients')+' <span class="muted">(first '+arr(r.preview).length+' of '+num(r.patients)+')</span></h3>'+table(r.headers,r.preview);
 arr(r.views).forEach(v=>{h+='<h3>'+esc(v.name)+' <span class="muted">('+plural(v.totalRows,'group')+', '+plural(v.patients,'patient')+(v.totalRows>arr(v.rows).length?'; first '+arr(v.rows).length+' shown':'')+')</span></h3>'+table(v.headers,v.rows,v.total,v.groupCount)});
 h+='<h3>What each rule did</h3><div class="scroll"><table class="grid"><thead><tr><th>#</th><th>Rule</th><th>What it does</th><th>Effect in this run</th></tr></thead><tbody>'+arr(r.effects).map(e=>'<tr><td class="n">'+(e.index||'')+'</td><td><span class="badge b-'+esc(String(e.type).split(' ')[0])+'">'+esc(e.type)+'</span></td><td>'+esc(e.summary)+(e.note?'<br><span class="muted">'+esc(e.note)+'</span>':'')+'</td><td>'+esc(e.effect)+(e.problems?'<div class="prob warn">'+esc(e.problems)+'</div>':'')+'</td></tr>').join('')+'</tbody></table></div>';
 h+='<h3>Sources</h3>'+table(['Source','Role','File','Sheet','Rows read','Removed (filters, unique)','Without key','Rows used','Patients'],arr(r.sources).map(s=>[s.source,s.role,s.file,s.sheet,num(s.rowsRead),num(s.filtered),num(s.noKey),num(s.rowsUsed),s.role==='patients'?num(s.patients):'']));
 if(arr(r.conflictSample).length)h+='<h3>Conflicts <span class="muted">(first '+arr(r.conflictSample).length+' of '+num(r.conflictRows)+'; all are on the Audit - Conflicts sheet)</span></h3>'+table(['Patient','Field','Chosen','From','Other values [source]'],arr(r.conflictSample).map(c=>[c.patient,c.field,c.chosen,c.from,c.others]));
 el('result').innerHTML=h;sync();
}
function historyHtml(list){
 if(!list.length)return '<span class="muted">None yet.</span>';
 return '<table class="grid"><thead><tr><th>Built</th><th>Config</th><th>Files</th><th class="n">Patients</th><th></th></tr></thead><tbody>'+list.map(r=>'<tr><td>'+esc(when(r.completedUtc||r.requestedUtc))+(current&&current.runId===r.runId?' <span class="muted">(shown above)</span>':'')+'</td><td>'+esc(r.configName)+(r.kind==='parity'?' <span class="badge b-key">side by side</span>'+(r.state==='Completed'?' <span class="'+(r.differentRows?'warn':'ok')+'">'+(r.differentRows?plural(r.differentRows,'patient')+' differ':'identical')+'</span>':''):'')+'</td><td>'+arr(r.assignments).map(a=>esc(a.source)+': '+esc(a.file)).join('<br>')+'</td><td class="n">'+num(r.patients)+'</td><td>'+(r.state==='Completed'?'<a href="#" data-view="'+esc(r.runId)+'">View</a>'+(r.parityUrl?' &middot; <a href="'+esc(r.parityUrl)+'">Comparison</a>':'')+(r.outputUrl?' &middot; <a href="'+esc(r.outputUrl)+'">Workbook</a>':' &middot; <span class="warn">workbook missing</span>'):'<span class="error">Failed: '+esc(r.errorSummary)+'</span>')+'</td></tr>').join('')+'</tbody></table>';
}
el('history').addEventListener('click',async e=>{const a=e.target.closest('a[data-view]');if(!a)return;e.preventDefault();if(busy())return;try{const o=await api('/api/combine/run?id='+encodeURIComponent(a.dataset.view));if(o.kind==='parity'){showParity(o);el('parityOut').scrollIntoView({behavior:'smooth'});return}showResult(o);el('history').innerHTML=historyHtml(history);el('result').scrollIntoView({behavior:'smooth'})}catch(err){el('result').innerHTML='<p class="error">'+esc(err.message)+'</p>'}});
let pollErrors=0;
function startPoll(id){if(timer)clearInterval(timer);pollErrors=0;timer=setInterval(async()=>{try{const o=await api('/api/combine/run?id='+encodeURIComponent(id));pollErrors=0;if(o.kind==='parity')showParity(o);else showResult(o);if(!busy()){clearInterval(timer);timer=null;await loadPage(false)}}catch(e){if(++pollErrors<5)return;clearInterval(timer);timer=null;el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>'}},1500)}
el('run').onclick=async()=>{if(el('run').disabled)return;el('run').disabled=true;try{const o=await post('/api/combine/run',{name:el('cfg').value,choices:choices});showResult(o);if(busy())startPoll(o.runId);else await loadPage(false)}catch(e){el('runNote').innerHTML='<span class="error">'+esc(e.message)+'</span>';sync()}};
async function loadPage(first){
 const m=await api('/api/combine/output');history=arr(m.history);el('folder').textContent=m.folder;
 canon=arr(m.canonicalSources);profileCount=m.profiles||0;paName=m.paConfig||paName;
 if(first){configs=arr(m.configs);const want=curConfig();el('cfg').innerHTML=configs.length?configs.map(n=>'<option value="'+esc(n)+'">'+esc(n)+'</option>').join(''):'<option value="">No configs yet</option>';el('cfg').value=configs.includes(want)?want:(configs[0]||'');await loadSetup();showResult(m.latest);showParity(m.latestParity);if(running(m.latest))startPoll(m.latest.runId);else if(running(m.latestParity))startPoll(m.latestParity.runId)}
 el('history').innerHTML=historyHtml(history);
}
loadPage(true).catch(e=>{el('result').innerHTML='<p class="error">'+esc(e.message)+'</p>'});
</script>
'@
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
 try{Invoke-ScheduledFeedSync}catch{Log 'FEED_SYNC' 'FAILED' $_.Exception.Message}
 # Like every step above, a failure here is logged; it must never stop the server (the heartbeat runs outside any request).
 if(!$pending){try{$pending=(@(Get-Jobs|Where-Object{[string](Get-P $_ 'state' '') -in @('Prepared','Queued','Starting','Running')}).Count -gt 0)}catch{$pending=$true;Log 'JOB_QUEUE' 'FAILED' $_.Exception.Message}}
 Set-KeepAwake $pending
}
function Invoke-AnalysisJob([string]$JobId){$path=Join-Path $script:Paths.State ('job-'+$JobId+'.json');$job=Json $path;if(!$job){throw 'Analysis job was not found.'};try{Set-JobProgress $job 5 'Starting analysis worker';Set-P $job 'startedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $job;$model=New-ProviderAnalysis $job;Set-JobProgress $job 72 'Rendering HTML dashboard';$slug=ConvertTo-ProviderSlug ([string]$job.displayName);$stamp=Get-Date -Format 'yyyyMMdd-HHmmss';$prefix=$slug+'__'+[string]$job.providerKey+'__analysis__';$base=$prefix+$stamp+'__'+$JobId.Substring(0,8);$stageHtml=Join-Path $script:Paths.Staging ($JobId+'.html');$stagePdf=Join-Path $script:Paths.Staging ($JobId+'.pdf');$jsonPath=Join-Path $script:Paths.State ('analysis-'+$JobId+'.json');Save-JsonAtomic $jsonPath $model;[IO.File]::WriteAllText($stageHtml,(ConvertTo-AnalysisHtml $model),(New-Object Text.UTF8Encoding($false)));Set-JobProgress $job 82 'Rendering PDF';Convert-HtmlPdf $stageHtml $stagePdf;Set-JobProgress $job 92 'Publishing current outputs';$archiveDir=Join-Path $script:Paths.ProvidersArchive $slug;if(!(Test-Path $archiveDir)){New-Item -ItemType Directory -Path $archiveDir -Force|Out-Null};$currentHtml=Join-Path $script:Paths.ProvidersCurrent ($base+'.html');$currentPdf=Join-Path $script:Paths.ProvidersCurrent ($base+'.pdf');$warnings=@();foreach($old in @(Get-ChildItem -LiteralPath $script:Paths.ProvidersCurrent -File -Filter ($prefix+'*') -ErrorAction SilentlyContinue)){if($old.BaseName -eq $base){continue};$dest=Join-Path $archiveDir $old.Name;if(Test-Path -LiteralPath $dest){$dest=Join-Path $archiveDir ($old.BaseName+'__'+$stamp+$old.Extension)};try{Move-Item -LiteralPath $old.FullName -Destination $dest -Force}catch{$warnings+=('Previous output '+$old.Name+' was left in providers-current (it is probably open in another program): '+$_.Exception.Message);Log 'ANALYSIS_ARCHIVE' 'WARN' $_.Exception.Message}};Move-Item -LiteralPath $stageHtml -Destination $currentHtml -Force;Move-Item -LiteralPath $stagePdf -Destination $currentPdf -Force;if($warnings.Count -gt 0){Set-P $job 'warnings' $warnings};$job.state='Completed';$job.percent=100;$job.stage='Analysis complete';Set-P $job 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Set-P $job 'htmlPath' $currentHtml;Set-P $job 'pdfPath' $currentPdf;Set-P $job 'htmlUrl' ('/html?jobId='+$JobId);Set-P $job 'pdfUrl' ('/pdf?jobId='+$JobId);Set-P $job 'flagUrl' ('/flag?jobId='+$JobId);Save-JsonAtomic $path $job;Log 'ANALYSIS_COMPLETED' 'OK' ('Job '+$JobId)}catch{$msg=$_.Exception.Message;if($msg -like 'Browser PDF rendering*' -and [int](Get-P $job 'attempts' 1) -lt $script:MaxJobAttempts){$job.state='Prepared';$job.percent=0;$job.stage=('Interrupted at 82% because '+$msg.TrimEnd('.')+'; queued to run again');Set-P $job 'errorSummary' '';Set-P $job 'workerPid' 0;Save-JsonAtomic $path $job;Log 'ANALYSIS_RETRY' 'WARN' ($JobId+' '+$msg)}else{$job.state='Failed';$job.stage='Analysis failed';Set-P $job 'errorSummary' $msg;Set-P $job 'completedUtc' ([DateTime]::UtcNow.ToString('o'));Save-JsonAtomic $path $job;Log 'ANALYSIS_FAILED' 'FAILED' ($JobId+' '+$msg)}}}
function Invoke-NextPendingJob{$jobs=@(Get-Jobs);if(@($jobs|Where-Object{$_.state -in @('Starting','Running')}).Count -gt 0){return};$j=@($jobs|Where-Object{$_.state -eq 'Prepared' -or $_.state -eq 'Queued'}|Sort-Object queuedUtc|Select-Object -First 1);if($j.Count -ne 1){return};$job=$j[0];$job.state='Starting';$job.percent=1;$attempt=1+[int](Get-P $job 'attempts' 0);Set-P $job 'attempts' $attempt;$job.stage=$(if($attempt -gt 1){'Launching background analysis worker (attempt '+$attempt+' of '+$script:MaxJobAttempts+')'}else{'Launching background analysis worker'});$jobPath=Join-Path $script:Paths.State ('job-'+$job.jobId+'.json');Save-JsonAtomic $jobPath $job;$exe=(Get-Process -Id $PID).Path;$out=Join-Path $script:Paths.Logs ('worker-'+$job.jobId+'.out.log');$err=Join-Path $script:Paths.Logs ('worker-'+$job.jobId+'.err.log');try{$quotedScript='"'+$PSCommandPath+'"';$p=Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$quotedScript,'-RunJobId',$job.jobId,'-NoBrowser') -RedirectStandardOutput $out -RedirectStandardError $err -WindowStyle Hidden -PassThru;Set-P $job 'workerPid' $p.Id;Save-JsonAtomic $jobPath $job}catch{$job.state='Failed';$job.stage='Worker launch failed';Set-P $job 'errorSummary' $_.Exception.Message;Save-JsonAtomic $jobPath $job}};function Invoke-PendingJobs{while(@(Get-Jobs|Where-Object{$_.state -eq 'Prepared' -or $_.state -eq 'Queued'}).Count -gt 0){Invoke-NextPendingJob;Start-Sleep -Milliseconds 500}}
function Send-JobOutput($Context,[string]$JobId,[string]$Type){if($JobId -notmatch '^[a-f0-9]{32}$'){throw 'Invalid job ID.'};$job=Json (Join-Path $script:Paths.State ('job-'+$JobId+'.json'));if(!$job){throw 'Job not found.'};$path=Resolve-JobOutputPath $job $Type;if(!$path){throw 'Output is not available; the file may have been moved or deleted.'};$bytes=[IO.File]::ReadAllBytes($path);$Context.Response.StatusCode=200;$Context.Response.ContentType=$(if($Type -eq 'pdf'){'application/pdf'}else{'text/html; charset=utf-8'});$Context.Response.Headers['Content-Disposition']='inline; filename="'+[IO.Path]::GetFileName($path)+'"';$Context.Response.ContentLength64=$bytes.Length;$Context.Response.Headers['Cache-Control']='no-store';$Context.Response.OutputStream.Write($bytes,0,$bytes.Length);$Context.Response.Close()}
function Initialize-AnalysisJobRecovery{$script:RunNext=$false;foreach($f in Get-ChildItem $script:Paths.State -Filter 'job-*.json' -File -ErrorAction SilentlyContinue){$j=Json $f.FullName;if(!$j){continue};$state=[string](Get-P $j 'state' '');$err=[string](Get-P $j 'errorSummary' '');if($state -in @('Starting','Running')){if(Test-WorkerAlive $j){$age=([DateTime]::UtcNow-(Get-Item -LiteralPath $f.FullName).LastWriteTimeUtc).TotalMinutes;if($age -lt $script:StartupStaleMinutes){continue};Stop-WorkerProcess ([int](Get-P $j 'workerPid' 0));Reset-InterruptedJob $j $f.FullName ('it had made no progress for '+[int]$age+' minutes when the server restarted')}else{Reset-InterruptedJob $j $f.FullName 'the server was restarted while the report was running'}}elseif($state -eq 'Failed' -and $err -like '*ConvertTo-NumberValue*'){$j.state='Prepared';$j.percent=0;$j.stage='Recovered after Draft 4.2 helper/queue repair';Set-P $j 'errorSummary' '';Save-JsonAtomic $f.FullName $j}}}
if($RunJobId){try{Initialize-AppFolders;Initialize-AppConfiguration;Test-ImportExcelModule;Invoke-AnalysisJob $RunJobId;exit 0}catch{Write-Error ($_.Exception.Message+' | '+$_.ScriptStackTrace);exit 1}};if($RunOverviewId){try{Initialize-AppFolders;Initialize-AppConfiguration;Test-ImportExcelModule;Invoke-OverviewRun $RunOverviewId;exit 0}catch{Write-Error ($_.Exception.Message+' | '+$_.ScriptStackTrace);exit 1}};if($RunNewListId){try{Initialize-AppFolders;Initialize-AppConfiguration;Test-ImportExcelModule;Invoke-NewListRun $RunNewListId;exit 0}catch{Write-Error ($_.Exception.Message+' | '+$_.ScriptStackTrace);exit 1}};if($RunPocId){try{Initialize-AppFolders;Initialize-AppConfiguration;Test-ImportExcelModule;Invoke-PocRun $RunPocId;exit 0}catch{Write-Error ($_.Exception.Message+' | '+$_.ScriptStackTrace);exit 1}};if($RunCombineId){try{Initialize-AppFolders;Initialize-AppConfiguration;Test-ImportExcelModule;Invoke-CmbRun $RunCombineId;exit 0}catch{Write-Error ($_.Exception.Message+' | '+$_.ScriptStackTrace);exit 1}};try{Initialize-AppFolders;Lock;Initialize-AppConfiguration;Test-ImportExcelModule;Initialize-AnalysisJobRecovery;Initialize-ProviderIndexes;Serve}catch{$detail=$_.Exception.Message+' | '+$_.ScriptStackTrace;try{Log 'APPLICATION' 'FAILED' $detail}catch{};Write-Error $detail;exit 1}finally{if($script:Listener){try{$script:Listener.Stop();$script:Listener.Close()}catch{}};if($script:Mutex){try{$script:Mutex.ReleaseMutex()}catch{};$script:Mutex.Dispose()};try{Log 'SERVER_STOP'}catch{}}