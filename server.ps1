# 주간업무보고 양식 통일 도구 - 로컬 서버
# 내 PC(127.0.0.1)에서만 열리며, 파일은 PC 밖으로 나가지 않습니다.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'convert.ps1')
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Web

$work = Join-Path $env:TEMP 'weekly-unify'
if (Test-Path $work) { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force (Join-Path $work 'in') | Out-Null
New-Item -ItemType Directory -Force (Join-Path $work 'out') | Out-Null

$jobs = @{}
$script:xl = $null
function Get-Excel {
  if ($script:xl) { try { $null = $script:xl.Version; return $script:xl } catch { $script:xl = $null } }
  $x = New-Object -ComObject Excel.Application
  $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
  try { $x.AskToUpdateLinks = $false } catch {}
  try { $x.EnableEvents = $false } catch {}
  try { $x.AutomationSecurity = 3 } catch {}
  $script:xl = $x
  return $x
}

# ---------- 최소 HTTP 처리 ----------
$utf8 = New-Object System.Text.UTF8Encoding $false
function Read-Request($ns) {
  $buf = New-Object byte[] 65536
  $ms = New-Object System.IO.MemoryStream
  $hEnd = -1
  while ($hEnd -lt 0) {
    $n = $ns.Read($buf, 0, $buf.Length)
    if ($n -le 0) { return $null }
    $ms.Write($buf, 0, $n)
    $hEnd = [System.Text.Encoding]::ASCII.GetString($ms.ToArray()).IndexOf("`r`n`r`n")
    if ($ms.Length -gt 1MB -and $hEnd -lt 0) { return $null }
  }
  $all = $ms.ToArray()
  $head = [System.Text.Encoding]::ASCII.GetString($all, 0, $hEnd)
  $lines = $head -split "`r`n"
  $first = $lines[0] -split ' '
  $headers = @{}
  foreach ($l in $lines[1..($lines.Count - 1)]) { $i = $l.IndexOf(':'); if ($i -gt 0) { $headers[$l.Substring(0, $i).Trim().ToLower()] = $l.Substring($i + 1).Trim() } }
  $len = 0; if ($headers['content-length']) { $len = [int64]$headers['content-length'] }
  $body = New-Object System.IO.MemoryStream
  $start = $hEnd + 4
  if ($all.Length -gt $start) { $body.Write($all, $start, $all.Length - $start) }
  while ($body.Length -lt $len) {
    $n = $ns.Read($buf, 0, [Math]::Min($buf.Length, $len - $body.Length))
    if ($n -le 0) { break }
    $body.Write($buf, 0, $n)
  }
  $url = $first[1]; $q = ''
  $qi = $url.IndexOf('?'); if ($qi -ge 0) { $q = $url.Substring($qi + 1); $url = $url.Substring(0, $qi) }
  return @{ method = $first[0]; path = $url; query = [System.Web.HttpUtility]::ParseQueryString($q); headers = $headers; body = $body.ToArray() }
}

function Send-Response($ns, [int]$code, [string]$ctype, [byte[]]$body, [string]$extra = '') {
  $reason = @{ 200 = 'OK'; 400 = 'Bad Request'; 404 = 'Not Found'; 500 = 'Internal Server Error' }[$code]
  $h = "HTTP/1.1 $code $reason`r`nContent-Type: $ctype`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n$extra`r`n"
  $hb = [System.Text.Encoding]::ASCII.GetBytes($h)
  $ns.Write($hb, 0, $hb.Length)
  if ($body.Length -gt 0) { $ns.Write($body, 0, $body.Length) }
  $ns.Flush()
}
function Send-Json($ns, $obj, [int]$code = 200) { Send-Response $ns $code 'application/json; charset=utf-8' ($utf8.GetBytes(($obj | ConvertTo-Json -Depth 6 -Compress))) }
function Get-Disposition([string]$name) {
  $enc = [Uri]::EscapeDataString($name)
  return "Content-Disposition: attachment; filename=`"download.xlsx`"; filename*=UTF-8''$enc`r`n"
}
function Get-SafeName([string]$name) {
  $n = $name
  foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $n = $n.Replace([string]$c, '_') }
  if (-not $n.ToLower().EndsWith('.xlsx')) { $n += '.xlsx' }
  return $n
}

# ---------- 요청 처리 ----------
function Handle($req, $ns) {
  switch -Regex ($req.path) {
    '^/$' {
      Send-Response $ns 200 'text/html; charset=utf-8' ([IO.File]::ReadAllBytes((Join-Path $here 'index.html')))
    }
    '^/api/ping$' { Send-Json $ns @{ ok = $true; date = (Get-Date -Format 'yyMMdd') } }
    '^/api/convert$' {
      $orig = [Uri]::UnescapeDataString([string]$req.headers['x-file-name'])
      if (-not $orig) { $orig = 'upload.xlsx' }
      $ext = [IO.Path]::GetExtension($orig).ToLower()
      if ($ext -notin @('.xlsx', '.xls', '.xlsm')) { Send-Json $ns @{ ok = $false; error = '엑셀 파일(.xlsx, .xls, .xlsm)만 넣을 수 있습니다.' }; return }
      $id = [guid]::NewGuid().ToString('N')
      $inDir = Join-Path $work "in\$id"; $outDir = Join-Path $work "out\$id"
      New-Item -ItemType Directory -Force $inDir, $outDir | Out-Null
      # 원본 확장자와 실제 형식이 달라도 Excel이 열 수 있도록 확장자는 그대로 둔다
      $inPath = Join-Path $inDir ("source" + $ext)
      [IO.File]::WriteAllBytes($inPath, $req.body)
      try {
        $res = Convert-WeeklyReport -Excel (Get-Excel) -InPath $inPath -OutDir $outDir -OrigName $orig -DateTag (Get-Date -Format 'yyMMdd')
        $jobs[$id] = $res.outPath
        Write-Host ("  완료: {0} -> {1}" -f $orig, $res.outName)
        Send-Json $ns @{ ok = $true; id = $id; original = $orig; outName = $res.outName; dept = $res.dept; deptHow = $res.deptHow; sheets = @($res.sheets); warnings = @($res.warnings) }
      } catch {
        Write-Host ("  실패: {0} ({1})" -f $orig, $_.Exception.Message) -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
        try { $null = $script:xl.Version } catch { $script:xl = $null }
        $where = '{0} {1}번째 줄' -f [IO.Path]::GetFileName($_.InvocationInfo.ScriptName), $_.InvocationInfo.ScriptLineNumber
        Send-Json $ns @{ ok = $false; original = $orig; error = ('{0} ({1})' -f $_.Exception.Message, $where) }
      }
    }
    '^/api/download$' {
      $id = $req.query['id']; $name = Get-SafeName ([string]$req.query['name'])
      if (-not $jobs.ContainsKey($id)) { Send-Json $ns @{ ok = $false; error = '파일을 찾을 수 없습니다.' } 404; return }
      Send-Response $ns 200 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ([IO.File]::ReadAllBytes($jobs[$id])) (Get-Disposition $name)
    }
    '^/api/zip$' {
      $items = $utf8.GetString($req.body) | ConvertFrom-Json
      $ms = New-Object System.IO.MemoryStream
      $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Create, $true)
      $used = @{}
      foreach ($it in @($items)) {
        if (-not $jobs.ContainsKey($it.id)) { continue }
        $n = Get-SafeName ([string]$it.name); $b = [IO.Path]::GetFileNameWithoutExtension($n); $k = 2
        while ($used.ContainsKey($n.ToLower())) { $n = "$b ($k).xlsx"; $k++ }
        $used[$n.ToLower()] = 1
        $e = $zip.CreateEntry($n)
        $s = $e.Open(); $bytes = [IO.File]::ReadAllBytes($jobs[$it.id]); $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
      }
      $zip.Dispose()
      $zname = "주간업무보고_양식통일본_{0}.zip" -f (Get-Date -Format 'yyMMdd')
      Send-Response $ns 200 'application/zip' $ms.ToArray() ("Content-Disposition: attachment; filename=`"download.zip`"; filename*=UTF-8''" + [Uri]::EscapeDataString($zname) + "`r`n")
    }
    '^/api/shutdown$' { Send-Json $ns @{ ok = $true }; $script:running = $false }
    default { Send-Response $ns 404 'text/plain; charset=utf-8' ($utf8.GetBytes('not found')) }
  }
}

# ---------- 시작 ----------
$listener = $null; $port = 0
foreach ($p in 8765..8775) {
  try { $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $p); $l.Start(); $listener = $l; $port = $p; break } catch {}
}
if (-not $listener) { Write-Host '사용 가능한 포트를 찾지 못했습니다.' -ForegroundColor Red; Read-Host '엔터를 누르면 닫힙니다'; exit 1 }

$url = "http://localhost:$port/"
Write-Host ''
Write-Host '  주간업무보고 양식 통일 도구가 실행 중입니다.' -ForegroundColor Cyan
Write-Host "  주소: $url"
Write-Host '  사용이 끝나면 화면의 [도구 종료] 버튼을 누르거나 이 창을 닫아 주세요.'
Write-Host ''
if ($env:WU_NOBROWSER -ne '1') { Start-Process $url }

$script:running = $true
try {
  while ($script:running) {
    $client = $listener.AcceptTcpClient()
    try {
      $ns = $client.GetStream(); $ns.ReadTimeout = 3000; $ns.WriteTimeout = 30000
      $req = Read-Request $ns
      if ($req) {
        try { Handle $req $ns }
        catch { try { Send-Json $ns @{ ok = $false; error = $_.Exception.Message } 500 } catch {} }
      }
    } catch {
      # 브라우저가 미리 열어둔 빈 연결 등은 무시
    } finally {
      $client.Close()
    }
  }
} finally {
  $listener.Stop()
  if ($script:xl) { try { $script:xl.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:xl) }
  Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
  Write-Host '  도구를 종료했습니다.'
}
