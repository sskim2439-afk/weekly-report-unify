# 주간업무보고 양식 통일 도구 - 자동 업데이트
# 실행할 때마다 GitHub에서 최신 파일을 받아온다. 실패하면 지금 버전으로 그대로 실행된다.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRaw = 'https://raw.githubusercontent.com/sskim2439-afk/weekly-report-unify/main'

# 각 파일이 정상인지 확인할 때 쓰는 표식 (받은 파일에 이 문구가 없으면 깨진 것으로 본다)
$files = [ordered]@{
  'convert.ps1' = 'function Convert-WeeklyReport'
  'server.ps1'  = 'function Send-Response'
  'index.html'  = '양식 통일 전'
  'update.ps1'  = '자동 업데이트'
  'README.md'   = '양식 통일 도구'
}

# 도구를 고치는 사람의 PC: 아직 GitHub에 올리지 않은 수정본이 덮어써지지 않도록 건너뛴다
if (Test-Path (Join-Path $here '업데이트안함.txt')) {
  Write-Host '  [업데이트] 업데이트안함.txt 가 있어 건너뜁니다.' -ForegroundColor DarkGray
  exit 0
}

Write-Host '  [업데이트] 최신 버전 확인 중...' -ForegroundColor Cyan
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  try { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch {}

  $tmp = Join-Path $env:TEMP ('weekly-unify-update-' + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Force $tmp | Out-Null
  $stamp = [DateTime]::Now.Ticks

  # 1) 전부 임시 폴더로 받고 검사한다
  foreach ($name in $files.Keys) {
    $url = '{0}/{1}?t={2}' -f $repoRaw, [Uri]::EscapeDataString($name), $stamp
    $dst = Join-Path $tmp $name
    Invoke-WebRequest -Uri $url -OutFile $dst -UseBasicParsing -TimeoutSec 15 -Headers @{ 'Cache-Control' = 'no-cache' }
    $text = [IO.File]::ReadAllText($dst, [Text.Encoding]::UTF8)
    if (-not $text.Contains($files[$name])) { throw "$name 내용이 올바르지 않습니다." }
    if ($name.EndsWith('.ps1')) {
      $errs = $null
      [void][System.Management.Automation.Language.Parser]::ParseFile($dst, [ref]$null, [ref]$errs)
      if ($errs.Count -gt 0) { throw "$name 에 문법 오류가 있습니다." }
    }
  }

  # 2) 모두 정상일 때만, 바뀐 파일만 교체한다
  $changed = @()
  foreach ($name in $files.Keys) {
    $new = [IO.File]::ReadAllBytes((Join-Path $tmp $name))
    $curPath = Join-Path $here $name
    $same = $false
    if (Test-Path $curPath) {
      $cur = [IO.File]::ReadAllBytes($curPath)
      $same = ($cur.Length -eq $new.Length) -and ([Convert]::ToBase64String($cur) -eq [Convert]::ToBase64String($new))
    }
    if (-not $same) { [IO.File]::WriteAllBytes($curPath, $new); $changed += $name }
  }
  Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

  if ($changed.Count -gt 0) { Write-Host ('  [업데이트] 새 버전으로 바꿨습니다: ' + ($changed -join ', ')) -ForegroundColor Green }
  else { Write-Host '  [업데이트] 이미 최신 버전입니다.' -ForegroundColor Green }
} catch {
  Write-Host ('  [업데이트] 확인하지 못해 지금 버전으로 실행합니다. ({0})' -f $_.Exception.Message) -ForegroundColor Yellow
  if ($tmp) { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}
exit 0
