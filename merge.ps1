# 주간업무보고 통합 (경영전략본부 통합 / 전사 통합)
# server.ps1에서 convert.ps1 다음에 dot-source 한다. Find-Layout · Get-Sections · Get-DeptName 등은 convert.ps1 것을 쓴다.

$script:HdrFill = 14348258   # 머리글 칸 연한 초록 (참고 파일과 같은 색)

# 파일 → 자리(슬롯) 판별. 위에 있을수록 먼저 검사한다 (해외경영관리팀을 경영관리팀보다 먼저 등)
$script:MergeSlots = @{
  strategy = [ordered]@{
    '해외경영관리팀' = @('해외'); 'ESG경영팀' = @('ESG'); 'CE팀' = @('CE'); '원가관리팀' = @('원가')
    '전략담당' = @('전략담당', '전략 담당'); '경영관리팀' = @('경영관리')
  }
  company  = [ordered]@{
    '디지털전략본부' = @('디지털'); '경영전략본부' = @('경영전략', 'CSO'); '경영지원본부' = @('경영지원'); '구매담당' = @('구매')
    '인사노무담당' = @('인사', '노무'); '생산본부' = @('생산'); '서비스본부' = @('서비스'); '설치본부' = @('설치')
    '안전지원실' = @('안전'); '한국사업본부' = @('한국', '영업'); '해외사업본부' = @('해외')
  }
}

# 시트 양식 (참고 파일 ★경영전략본부 / ★전사 _주간업무보고_26.11.23 과 같게)
$script:SheetSpecs = @{
  s1 = @{ name = '경영관리담당 (금주 팀별)'; title = '현대 E/L 주간업무보고 (경영관리담당, 경영전략본부 직속 금주 팀별)'; titleSize = 15
          widths = @(0.88, 9.9, 80.2, 80.2, 0, 1.2); lab = 2; c1 = 3; c2 = 4; note = 5; lastCol = 6; firstCol = 1
          titleH = 26.3; gapH = 30; hdrH = 25.2; bodyH = 21.8; labelSize = 11; margins = @(17, 17, 53.9, 53.9); center = $true }
  s2 = @{ name = '전략담당 (금주 팀별)'; title = '현대 E/L 주간업무보고 (전략담당, 경영전략본부 직속 금주 팀별)'; titleSize = 15
          widths = @(1.2, 9.9, 80.2, 80.2, 0, 1.2); lab = 2; c1 = 3; c2 = 4; note = 5; lastCol = 6; firstCol = 1
          titleH = 26.3; gapH = 30; hdrH = 25.2; bodyH = 23.3; labelSize = 10; margins = @(17, 17, 53.9, 53.9); center = $true }
  s3 = @{ name = '금주 요약본'; title = '현대 E/L 주간업무보고 (경영전략본부 금주 요약본)'; titleSize = 15
          widths = @(4.5, 1.2, 9.9, 80.2, 80.2, 0, 1.2); lab = 3; c1 = 4; c2 = 5; note = 6; lastCol = 7; firstCol = 2
          titleH = 26.3; gapH = 30; hdrH = 25.2; bodyH = 21; labelSize = 11; margins = @(17, 17, 53.9, 53.9); center = $true }
  co = @{ name = '전사'; title = '현대 E/L 주간업무보고'; titleSize = 18
          widths = @(1.2, 11.8, 85, 85, 1.5); lab = 2; c1 = 3; c2 = 4; note = 0; lastCol = 5; firstCol = 1
          titleH = 26.4; gapH = 16.2; hdrH = 16.2; bodyH = 16.2; labelSize = 11; margins = @(8.5, 8.5, 45.4, 22.7); center = $false }
}

# 전사 시트의 페이지 구성 (참고 파일 순서 그대로). max = 그 페이지 묶음이 넘으면 안 되는 쪽수 (0 = 제한 없음)
$script:CompanyPages = @(
  @{ blocks = @(@{ label = '안전지원실'; parts = @('안전지원실') }, @{ label = '한국사업본부'; parts = @('한국사업본부') }); max = 0 },
  @{ blocks = @(@{ label = '해외사업본부'; parts = @('해외사업본부') }, @{ label = '서비스사업본부'; parts = @('서비스본부') }); max = 0 },
  @{ blocks = @(@{ label = '경영전략본부'; parts = @('경영전략본부') }); max = 0 },
  # 경영지원본부: [기획재무] → [구매] → [영업관리/커뮤니케이션] → [준법] → [인사/노무] 순서가 되도록
  # 경영지원본부 파일 안의 '기획재무' 덩어리 뒤에 구매, '준법' 덩어리 뒤에 인사노무를 끼워 넣는다. 2쪽을 넘으면 덜 중요한 내용을 표 밖으로
  @{ blocks = @(@{ label = '경영지원본부'; parts = @('경영지원본부')
                   inserts = @(@{ part = '구매담당'; bracket = '[구매담당]'; after = '기획재무'; atEnd = $false }
                               @{ part = '인사노무담당'; bracket = '[인사/노무담당]'; after = '준법'; atEnd = $true }) }); max = 2 },
  @{ blocks = @(@{ label = '생산본부'; parts = @('생산본부') }, @{ label = '설치본부'; parts = @('설치본부') }); max = 0 }
)
$script:StrategyOrder = @('경영관리팀', '원가관리팀', 'CE팀', 'ESG경영팀', '해외경영관리팀')

# ---------- 원본 읽기 ----------

# 내용 칸이 있는 보이는 시트 중 하나를 고른다. $prefer(정규식)에 맞는 이름이 있으면 그 시트
function Find-ReportSheet($wb, [string]$prefer) {
  # 열려 있던 시트를 먼저 본다: 지난 주차 시트가 수십 개인 파일도 빨리 끝나도록
  if (-not $prefer) { $act = $wb.ActiveSheet; if ($act.Visible -eq -1) { $lay = Find-Layout $act; if ($lay) { return @{ ws = $act; lay = $lay } } } }
  else { foreach ($ws in $wb.Worksheets) { if ($ws.Visible -eq -1 -and $ws.Name -match $prefer) { $lay = Find-Layout $ws; if ($lay) { return @{ ws = $ws; lay = $lay } } } } }
  $cands = @()
  foreach ($ws in $wb.Worksheets) { if ($ws.Visible -eq -1) { $lay = Find-Layout $ws; if ($lay) { $cands += @{ ws = $ws; lay = $lay } } } }
  if ($cands.Count -eq 0) { return $null }
  if ($prefer) { $p = @($cands | Where-Object { $_.ws.Name -match $prefer }); if ($p.Count) { return $p[0] } }
  $act = $wb.ActiveSheet.Name
  $a = @($cands | Where-Object { $_.ws.Name -eq $act }); if ($a.Count) { return $a[0] }
  return $cands[0]
}

# 파일이 어느 자리인지 판별 (부서명 → 파일명 순). $slimPath가 있으면 통합에 쓸 시트만 남긴 가벼운 사본을 저장한다
function Get-MergeSlot($Excel, [string]$path, [string]$origName, [string]$kind, [string]$slimPath) {
  $wb = $Excel.Workbooks.Open($path, 0, $true)
  try {
    $rs = Find-ReportSheet $wb $null
    if (-not $rs) { throw "'금주 / 차주' 또는 '전주 / 금주' 머리글이 있는 시트를 찾지 못했습니다." }
    $dept = (Get-DeptName $rs.ws $rs.lay (Get-Sections $rs.ws $rs.lay) $origName).name
    $slot = $null; $slots = $script:MergeSlots[$kind]
    foreach ($text in @($dept, [IO.Path]::GetFileNameWithoutExtension($origName))) {
      foreach ($k in $slots.Keys) { foreach ($kw in $slots[$k]) { if (-not $slot -and $text -match [regex]::Escape($kw)) { $slot = $k } } }
    }
    if ($slot -and $slimPath) {
      # 통합 때 다시 볼 시트: 기본 시트 + 요약본(전사) / 팀별·요약본(전략담당)
      $keep = @($rs.ws.Name)
      foreach ($pat in @('요약') + @($(if ($slot -eq '전략담당') { '팀별' }))) { $x = Find-ReportSheet $wb $pat; if ($x -and $keep -notcontains $x.ws.Name) { $keep += $x.ws.Name } }
      $sheets = @($keep | ForEach-Object { $wb.Worksheets.Item($_) })
      $wb.Worksheets.Item($keep[0]).Copy()   # 새 통합 문서로 복사
      $slim = $Excel.ActiveWorkbook
      for ($i = 1; $i -lt $sheets.Count; $i++) { $sheets[$i].Copy([Type]::Missing, $slim.Worksheets.Item($slim.Worksheets.Count)) }
      [void]$slim.Worksheets.Item(1).Activate()
      [void]$slim.SaveAs($slimPath, 51); [void]$slim.Close($false)
    }
    return @{ slot = $slot; dept = $dept }
  } finally { $wb.Close($false) }
}

# 시트의 구역들: 각 구역의 내용 줄 범위(앞뒤 빈 줄 제외)와 구분 칸 이름
function Read-Parts($rs) {
  $ws = $rs.ws; $lay = $rs.lay; $res = @()
  foreach ($sec in (Get-Sections $ws $lay)) {
    $t = Read-ColText $rs $sec[0] $sec[1]
    $r0 = 0; $r1 = 0
    for ($i = 0; $i -lt $t.Count; $i++) { if ($t[$i][0] -or $t[$i][1]) { if (-not $r0) { $r0 = $sec[0] + $i }; $r1 = $sec[0] + $i } }
    if ($r0) { $res += @{ label = (Clean-Label ([string]$sec[2])); r0 = $r0; r1 = $r1 } }
  }
  return , $res
}

# 두 내용 칸의 글자를 한 번에 읽는다 (칸마다 COM 호출하면 느림). 반환: 줄마다 @(왼쪽, 오른쪽)
function Read-ColText($rs, [int]$r0, [int]$r1) {
  $out = New-Object System.Collections.ArrayList
  $vals = @(); foreach ($c in $rs.lay.cols) { $vals += , $rs.ws.Range($rs.ws.Cells($r0, $c), $rs.ws.Cells($r1, $c)).Value2 }
  for ($i = 1; $i -le $r1 - $r0 + 1; $i++) {
    $pair = @(); foreach ($v in $vals) { $pair += $(if ($r1 -gt $r0) { ([string]$v[$i, 1]).Trim() } else { ([string]$v).Trim() }) }
    [void]$out.Add($pair)
  }
  return , $out
}

# 머리글 바로 아래 날짜 줄 '(11/16 ~ 11/22)' 두 개
function Read-WeekDates($rs) {
  $r = $rs.lay.headerRow + 1; $d = @()
  foreach ($c in $rs.lay.cols) { $v = ([string]$rs.ws.Cells($r, $c).Value2).Trim(); $d += $(if ($v -match '\d{1,2}\s*/\s*\d{1,2}') { $v } else { '' }) }
  return , $d
}

# ---------- 줄 계획 (메모리) ----------
# 한 줄 = @{ k='src'|'text'|'blank'; ws; r; c1; c2; text; band }  band가 같은 줄은 한 덩어리(주제)로 같이 움직인다

function Add-SourceRows([System.Collections.ArrayList]$plan, $rs, $r0, $r1, [ref]$bandSeq) {
  $ws = $rs.ws; $c1 = $rs.lay.cols[0]; $c2 = $rs.lay.cols[1]
  $band = ++$bandSeq.Value; $prevBlank = $true
  $txt = Read-ColText $rs $r0 $r1
  for ($r = $r0; $r -le $r1; $r++) {
    $t1 = $txt[$r - $r0][0]; $t2 = $txt[$r - $r0][1]
    # 원본의 빈 줄도 원본 위치를 기억해 둔다: 이어진 줄을 한 번에 복사할 수 있도록
    if (-not $t1 -and -not $t2) { [void]$plan.Add(@{ k = 'blank'; ws = $ws; r = $r; c1 = $c1; c2 = $c2 }); $prevBlank = $true; continue }
    # 원본 안의 [대괄호 제목] 줄은 혼자 한 덩어리 (표 밖으로 빼지 않음)
    $isBr = ($t1 -match '^\[' -or $t2 -match '^\[') -and ($t1 -eq '' -or $t1 -match '^\[') -and ($t2 -eq '' -or $t2 -match '^\[')
    if ($isBr) { [void]$plan.Add(@{ k = 'src'; ws = $ws; r = $r; c1 = $c1; c2 = $c2; band = 0; t = "$t1 $t2" }); $prevBlank = $true; continue }
    if ($prevBlank) { $band = ++$bandSeq.Value }
    [void]$plan.Add(@{ k = 'src'; ws = $ws; r = $r; c1 = $c1; c2 = $c2; band = $band; t = "$t1 $t2" })
    $prevBlank = $false
  }
}

# 한 블록(구분 칸 이름 하나)에 들어갈 줄들: 위아래 빈 줄 + 원본들. 원본 구역이 여러 개면 [구역 이름]을 넣어 구분
function New-BlockPlan($sources, [ref]$bandSeq) {
  $plan = New-Object System.Collections.ArrayList
  [void]$plan.Add(@{ k = 'blank' })
  $first = $true
  foreach ($src in $sources) {
    if (-not $first) { [void]$plan.Add(@{ k = 'blank' }) }
    $first = $false
    if ($src.bracket) { [void]$plan.Add(@{ k = 'text'; text = $src.bracket; band = 0 }) }
    $parts = Read-Parts $src.rs
    $labels = @($parts | ForEach-Object { $_.label } | Where-Object { $_ } | Select-Object -Unique)
    $pi = 0
    foreach ($p in $parts) {
      if ($pi -gt 0) { [void]$plan.Add(@{ k = 'blank' }) }
      $startsBr = (([string]$src.rs.ws.Cells($p.r0, $src.rs.lay.cols[0]).Value2).Trim() -match '^\[')
      if (-not $src.bracket -and $labels.Count -gt 1 -and $p.label -and -not $startsBr) { [void]$plan.Add(@{ k = 'text'; text = "[$($p.label)]"; band = 0 }) }
      Add-SourceRows $plan $src.rs $p.r0 $p.r1 $bandSeq
      $pi++
    }
  }
  [void]$plan.Add(@{ k = 'blank' })
  Remove-EmptyHeads $plan
  return , $plan
}

# 덜 중요해 보이는 덩어리 하나를 고른다: '기타' 제목 → (지속)/(계속) → 그 밖, 같은 점수면 아래쪽부터
function Select-DropBand($plan) {
  $bands = @{}
  for ($i = 0; $i -lt $plan.Count; $i++) {
    $x = $plan[$i]; if ($x.k -ne 'src' -or -not $x.band) { continue }
    if (-not $bands.ContainsKey($x.band)) { $bands[$x.band] = @{ id = $x.band; first = $i; last = $i; title = $x.t; all = '' } }
    $bands[$x.band].last = $i; $bands[$x.band].all += ' ' + $x.t
  }
  if ($bands.Count -eq 0) { return $null }
  $scored = $bands.Values | ForEach-Object {
    $s = 2; if ($_.title -match '기타') { $s = 0 } elseif ($_.all -match '지속|계속') { $s = 1 }
    [pscustomobject]@{ b = $_; s = $s; pos = $_.first }
  }
  return ($scored | Sort-Object s, @{ e = 'pos'; Descending = $true } | Select-Object -First 1).b
}

# 계획에서 덩어리를 빼고(뒤 빈 줄 하나도 같이) 빠진 줄 목록을 돌려준다
function Remove-Band([System.Collections.ArrayList]$plan, $band) {
  $out = @(); for ($i = $band.first; $i -le $band.last; $i++) { $out += $plan[$band.first] ; $plan.RemoveAt($band.first) }
  if ($band.first -lt $plan.Count -and $plan[$band.first].k -eq 'blank' -and $band.first -gt 0 -and $plan[$band.first - 1].k -eq 'blank') { $plan.RemoveAt($band.first) }
  return , $out
}

# 다른 파일의 내용을 [제목]과 함께 줄 목록으로 만든다 (블록 안에 끼워 넣을 때 사용)
function New-SubRows($rs, [string]$bracket, [ref]$bandSeq) {
  $l = New-Object System.Collections.ArrayList
  [void]$l.Add(@{ k = 'text'; text = $bracket; band = 0 })
  $first = $true
  foreach ($p in (Read-Parts $rs)) { if (-not $first) { [void]$l.Add(@{ k = 'blank' }) }; $first = $false; Add-SourceRows $l $rs $p.r0 $p.r1 $bandSeq }
  return , $l
}

# $rows를 $plan 안에서 '$after' 글자가 든 [제목] 덩어리 바로 뒤(다음 [제목] 앞)에 끼워 넣는다.
# 그런 제목이 없으면 $atEnd에 따라 맨 끝 / 맨 앞에 넣는다
function Add-IntoPlan([System.Collections.ArrayList]$plan, $rows, [string]$after, [bool]$atEnd) {
  $isHead = { param($x) $x.k -eq 'text' -or ($x.k -eq 'src' -and $x.band -eq 0) }
  $headText = { param($x) [string]$(if ($x.k -eq 'text') { $x.text } else { $x.t }) }
  $pos = -1
  for ($i = 0; $i -lt $plan.Count; $i++) { if ((& $isHead $plan[$i]) -and (& $headText $plan[$i]) -match $after) { $pos = $i; break } }
  $ins = New-Object System.Collections.ArrayList
  if ($pos -ge 0) {
    $j = $pos + 1; while ($j -lt $plan.Count -and -not (& $isHead $plan[$j])) { $j++ }
    if ($j -lt $plan.Count) { $ins.AddRange($rows); [void]$ins.Add(@{ k = 'blank' }) }          # 다음 [제목] 앞에: 내용 + 빈 줄
    else { $j = $plan.Count - 1; [void]$ins.Add(@{ k = 'blank' }); $ins.AddRange($rows) }      # 맨 끝 빈 줄 앞에: 빈 줄 + 내용
  } elseif ($atEnd) { $j = $plan.Count - 1; [void]$ins.Add(@{ k = 'blank' }); $ins.AddRange($rows) }
  else { $j = 1; $ins.AddRange($rows); [void]$ins.Add(@{ k = 'blank' }) }                         # 맨 앞 빈 줄 다음에
  $plan.InsertRange($j, $ins)
  Remove-EmptyHeads $plan
}

# 아래 내용이 모두 빠져 제목만 남은 [대괄호] 줄과, 겹친 빈 줄을 지운다
function Remove-EmptyHeads([System.Collections.ArrayList]$plan) {
  $isHead = { param($x) $x.k -eq 'text' -or ($x.k -eq 'src' -and $x.band -eq 0) }
  for ($i = $plan.Count - 1; $i -ge 0; $i--) {
    if (-not (& $isHead $plan[$i])) { continue }
    $j = $i + 1; while ($j -lt $plan.Count -and $plan[$j].k -eq 'blank') { $j++ }
    if ($j -ge $plan.Count -or (& $isHead $plan[$j])) { $plan.RemoveAt($i) }
  }
  for ($i = $plan.Count - 1; $i -ge 1; $i--) { if ($plan[$i].k -eq 'blank' -and $plan[$i - 1].k -eq 'blank') { $plan.RemoveAt($i) } }
}

# ---------- 시트 쓰기 ----------

function New-ReportSheet($wb, $sp, [bool]$useFirst) {
  $ws = $(if ($useFirst) { $wb.Worksheets.Item(1) } else { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) })
  $ws.Name = $sp.name
  $ws.Cells.Font.Name = '네오현대 R'
  for ($i = 0; $i -lt $sp.widths.Count; $i++) { $ws.Columns($i + 1).ColumnWidth = [double]$sp.widths[$i] }
  $ws.Cells.VerticalAlignment = -4108
  $ws.Rows(1).RowHeight = [double]16.2
  return $ws
}

# 머리글 묶음 5줄: [빈 줄, 제목, 빈 줄, 구분/금주/차주(/비고), 팀/날짜]
function Write-PageHeader($ws, $sp, [int]$row, $dates) {
  $ws.Rows($row).RowHeight = [double]16.2
  $t = $row + 1; $h1 = $row + 3; $h2 = $row + 4
  $ws.Rows($t).RowHeight = [double]$sp.titleH; $ws.Rows($row + 2).RowHeight = [double]$sp.gapH
  $ws.Rows($h1).RowHeight = [double]$sp.hdrH; $ws.Rows($h2).RowHeight = [double]$sp.hdrH
  $endCol = $(if ($sp.note) { $sp.note } else { $sp.c2 })
  $tr = $ws.Range($ws.Cells($t, $sp.lab), $ws.Cells($t, $endCol)); [void]$tr.Merge(); $tr.Value2 = $sp.title
  $tr.Font.Name = '네오현대 EBK'; $tr.Font.Size = [double]$sp.titleSize; $tr.HorizontalAlignment = -4108
  $ws.Cells($h1, $sp.lab).Value2 = '구 분'; $ws.Cells($h2, $sp.lab).Value2 = '팀'
  $ws.Cells($h1, $sp.c1).Value2 = '금주'; $ws.Cells($h1, $sp.c2).Value2 = '차주'
  $ws.Cells($h2, $sp.c1).Value2 = $dates[0]; $ws.Cells($h2, $sp.c2).Value2 = $dates[1]
  $hr = $ws.Range($ws.Cells($h1, $sp.lab), $ws.Cells($h2, $endCol))
  $hr.Interior.Color = [double]$script:HdrFill; $hr.Font.Name = '네오현대 B'; $hr.Font.Size = [double]10; $hr.HorizontalAlignment = -4108
  if ($sp.note) { $nr = $ws.Range($ws.Cells($h1, $sp.note), $ws.Cells($h2, $sp.note)); [void]$nr.Merge(); $nr.Value2 = '비고'; $nr.Font.Name = '네오현대 R' }
  foreach ($c in @($sp.lab, $sp.c1, $sp.c2) + @($(if ($sp.note) { $sp.note }))) { $ws.Range($ws.Cells($h1, $c), $ws.Cells($h2, $c)).BorderAround(1, 2) | Out-Null }
  $mid = $ws.Range($ws.Cells($h1, $sp.lab), $ws.Cells($h1, $sp.lab)).Borders.Item(9); $mid.LineStyle = 1; $mid.Weight = 1
}

# 한 페이지에 들어가는 본문 줄 수를 엑셀에게 직접 물어 본다 (본문 줄 높이가 모두 같으므로 줄 수로 계산 가능)
function Measure-PageRows($ws, $sp) {
  Set-PrintSetup $ws $sp 400
  $ws.Range($ws.Rows(1), $ws.Rows(400)).RowHeight = [double]$sp.bodyH
  $ws.Cells(400, $sp.lab).Value2 = '.'
  # 새 통합 문서는 페이지 나누기 미리보기로 바꿔야 엑셀이 쪽 나눔을 계산한다
  [void]$ws.Activate(); $win = $ws.Application.ActiveWindow; $view = $win.View; $win.View = 2
  $br = @(); foreach ($b in $ws.HPageBreaks) { $br += [int]$b.Location.Row }
  $win.View = $view
  [void]$ws.Cells(400, $sp.lab).ClearContents(); $ws.Rows(1).RowHeight = [double]16.2
  if ($br.Count -lt 2) { return 0 }
  return ($br[1] - $br[0])
}

function Set-PrintSetup($ws, $sp, [int]$lastRow) {
  $app = $ws.Application
  try { $app.PrintCommunication = $false } catch {}
  $ps = $ws.PageSetup
  $ps.PrintArea = $ws.Range($ws.Cells(1, $sp.firstCol), $ws.Cells($lastRow, $sp.lastCol)).Address()
  $ps.Orientation = 1; $ps.PaperSize = 9; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
  $ps.LeftMargin = [double]$sp.margins[0]; $ps.RightMargin = [double]$sp.margins[1]; $ps.TopMargin = [double]$sp.margins[2]; $ps.BottomMargin = [double]$sp.margins[3]
  $ps.CenterHorizontally = $sp.center
  try { $app.PrintCommunication = $true } catch {}
}

# 계획한 줄들을 시트에 쓴다. $pager = @{ used; cap(이번 쪽 본문 줄 수); next(다음 쪽부터); header(쪽마다 머리글 반복) }
# 쪽이 바뀌면 쪽 나눔을 넣고, 구분 칸도 쪽마다 따로 합치도록 표시를 바꾼다
function Write-Plan($ws, $sp, $plan, [int]$row, [string]$label, $pager, $dates, [ref]$markSeq) {
  $mark = "L|$(++$markSeq.Value)|$label"
  $run = $null   # 원본의 이어진 줄들은 모아서 한 번에 복사한다 (줄마다 복사하면 매우 느림)
  $segStart = $row
  foreach ($x in $plan) {
    if ($pager.cap -gt 0 -and $pager.used -ge $pager.cap) {
      Copy-Run $run $ws $sp; $run = $null
      Set-RowMarks $ws $sp $segStart ($row - 1) $mark
      [void]$ws.HPageBreaks.Add($ws.Cells($row, 1))
      if ($pager.header) { Write-PageHeader $ws $sp $row $dates; $row += 5 }
      $pager.used = 0; $pager.cap = $pager.next
      $mark = "L|$(++$markSeq.Value)|$label"; $segStart = $row
    }
    if ($x.ws) {
      if ($run -and $run.ws -eq $x.ws -and $run.c1 -eq $x.c1 -and $x.r -eq $run.r1 + 1) { $run.r1 = $x.r }
      else { Copy-Run $run $ws $sp; $run = @{ ws = $x.ws; c1 = $x.c1; c2 = $x.c2; r0 = $x.r; r1 = $x.r; dst = $row } }
    } else {
      Copy-Run $run $ws $sp; $run = $null
      if ($x.k -eq 'text') {
        foreach ($c in $sp.c1, $sp.c2) { $cell = $ws.Cells($row, $c); $cell.Value2 = $x.text; $cell.Font.Name = '네오현대 B'; $cell.Font.Size = [double]11 }
      }
    }
    $row++; $pager.used++
  }
  Copy-Run $run $ws $sp
  Set-RowMarks $ws $sp $segStart ($row - 1) $mark
  return $row
}

function Copy-Run($run, $ws, $sp) {
  if (-not $run) { return }
  $s = $run.ws
  if ($run.c2 -eq $run.c1 + 1) { [void]$s.Range($s.Cells($run.r0, $run.c1), $s.Cells($run.r1, $run.c2)).Copy($ws.Cells($run.dst, $sp.c1)) }
  else {
    [void]$s.Range($s.Cells($run.r0, $run.c1), $s.Cells($run.r1, $run.c1)).Copy($ws.Cells($run.dst, $sp.c1))
    [void]$s.Range($s.Cells($run.r0, $run.c2), $s.Cells($run.r1, $run.c2)).Copy($ws.Cells($run.dst, $sp.c2))
  }
}

# 칸 A에 구분 표시를 한 번에 적고 줄 높이를 맞춘다
function Set-RowMarks($ws, $sp, [int]$r0, [int]$r1, [string]$mark) {
  if ($r1 -lt $r0) { return }
  $ws.Range($ws.Cells($r0, 1), $ws.Cells($r1, 1)).Value2 = $mark
  $ws.Range($ws.Rows($r0), $ws.Rows($r1)).RowHeight = [double]$sp.bodyH
}

# 칸 A에 적어 둔 표시를 따라 구분 칸을 합치고 테두리를 그린다
function Complete-Labels($ws, $sp, [int]$first, [int]$last) {
  $va = $ws.Range($ws.Cells($first, 1), $ws.Cells($last, 1)).Value2
  $markAt = { param($row) [string]$va[($row - $first + 1), 1] }
  $r = $first
  while ($r -le $last) {
    $m = & $markAt $r
    if (-not $m.StartsWith('L|')) { $r++; continue }
    $e = $r; while ($e + 1 -le $last -and (& $markAt ($e + 1)) -eq $m) { $e++ }
    # 원본에서 딸려 온 테두리는 지우고 이 양식의 테두리만 그린다
    $ws.Range($ws.Cells($r, $sp.lab), $ws.Cells($e, $(if ($sp.note) { $sp.note } else { $sp.c2 }))).Borders.LineStyle = -4142
    $lr = $ws.Range($ws.Cells($r, $sp.lab), $ws.Cells($e, $sp.lab)); [void]$lr.Merge()
    $lr.Value2 = ($m -split '\|', 3)[2]; $lr.Font.Name = '네오현대 B'; $lr.Font.Size = [double]$sp.labelSize
    $lr.HorizontalAlignment = -4108; $lr.VerticalAlignment = -4108; $lr.WrapText = $true
    foreach ($c in @($sp.lab, $sp.c1, $sp.c2) + @($(if ($sp.note) { $sp.note }))) { $ws.Range($ws.Cells($r, $c), $ws.Cells($e, $c)).BorderAround(1, 2) | Out-Null }
    $r = $e + 1
  }
  [void]$ws.Range($ws.Cells($first, 1), $ws.Cells($last, 1)).ClearContents()
}

# 표 밖으로 뺀 줄들을 표 오른쪽(인쇄 범위 밖)에 붙인다
function Write-Dropped($ws, $sp, [int]$row, $dropped) {
  if (-not $dropped -or $dropped.Count -eq 0) { return }
  $oc = $sp.lastCol + 2
  $ws.Columns($oc).ColumnWidth = [double]85; $ws.Columns($oc + 1).ColumnWidth = [double]85
  $h = $ws.Range($ws.Cells($row, $oc), $ws.Cells($row, $oc + 1)); [void]$h.Merge()
  $h.Value2 = '※ 2쪽을 넘어 표 밖으로 뺀 내용 (왼쪽 금주 | 오른쪽 차주) - 검토 후 필요하면 표 안으로 옮겨 주세요'
  $h.Interior.Color = [double]65535; $h.Font.Name = '네오현대 B'; $h.Font.Size = [double]11
  $r = $row + 1
  foreach ($band in $dropped) {
    # 한 덩어리는 원본에서 이어진 줄이므로 한 번에 복사한다
    $x = $band[0]; $n = $band.Count
    [void]$x.ws.Range($x.ws.Cells($x.r, $x.c1), $x.ws.Cells($x.r + $n - 1, $x.c2)).Copy($ws.Cells($r, $oc))
    $r += $n + 1
  }
}

# ---------- 경영전략본부 통합 ----------

function Merge-Strategy($Excel, $files, [string]$outDir, [string]$dateTag) {
  $warn = New-Object System.Collections.ArrayList
  $open = @{}
  try {
    foreach ($k in $files.Keys) { $open[$k] = $Excel.Workbooks.Open($files[$k], 0, $true) }
    $team = @{}; foreach ($k in $script:StrategyOrder) { $team[$k] = Find-ReportSheet $open[$k] $null }
    $sTeam = Find-ReportSheet $open['전략담당'] '팀별'
    $sSum = Find-ReportSheet $open['전략담당'] '요약'
    $dates = Read-WeekDates $team['경영관리팀']
    $wb = $Excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { [void]$wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $bandSeq = 0; $markSeq = 0

    # 시트1: 경영관리담당 팀별 = 5개 팀 전체
    $sp = $script:SheetSpecs.s1; $ws = New-ReportSheet $wb $sp $true
    Write-PageHeader $ws $sp 2 $dates; $row = 7; $noPage = @{ used = 0; cap = 0; next = 0; header = $false }
    foreach ($k in $script:StrategyOrder) {
      $plan = New-BlockPlan @(@{ rs = $team[$k] }) ([ref]$bandSeq)
      $row = Write-Plan $ws $sp $plan $row $k $noPage $dates ([ref]$markSeq)
    }
    Complete-Labels $ws $sp 7 ($row - 1); Set-PrintSetup $ws $sp ($row - 1)

    # 시트2: 전략담당 팀별 = 전략담당 파일의 팀별 시트 (팀마다 구분 칸)
    $sp = $script:SheetSpecs.s2; $ws = New-ReportSheet $wb $sp $false
    Write-PageHeader $ws $sp 2 $dates; $row = 7
    foreach ($p in (Read-Parts $sTeam)) {
      $plan = New-Object System.Collections.ArrayList; [void]$plan.Add(@{ k = 'blank' })
      Add-SourceRows $plan $sTeam $p.r0 $p.r1 ([ref]$bandSeq); [void]$plan.Add(@{ k = 'blank' })
      $row = Write-Plan $ws $sp $plan $row $p.label $noPage $dates ([ref]$markSeq)
    }
    Complete-Labels $ws $sp 7 ($row - 1); Set-PrintSetup $ws $sp ($row - 1)

    # 시트3: 금주 요약본 = [경영관리담당] 5개 팀 + [전략담당] 요약본. 2쪽을 넘으면 덜 중요한 내용을 표 밖으로
    $sp = $script:SheetSpecs.s3; $ws = New-ReportSheet $wb $sp $false
    $pageRows = Measure-PageRows $ws $sp
    Write-PageHeader $ws $sp 2 $dates
    $hdrRows = [Math]::Ceiling((16.2 * 2 + $sp.titleH + $sp.gapH + $sp.hdrH * 2) / $sp.bodyH)
    $cap1 = $pageRows - $hdrRows - 1; $capN = $pageRows - 1
    $srcs = @(@{ rs = $team[$script:StrategyOrder[0]]; bracket = '[경영관리담당]' }) + @($script:StrategyOrder | Select-Object -Skip 1 | ForEach-Object { @{ rs = $team[$_] } })
    $srcs += @{ rs = $(if ($sSum) { $sSum } else { $sTeam }); bracket = '[전략담당]' }
    $plan = New-BlockPlan $srcs ([ref]$bandSeq)
    $dropped = @()
    if ($pageRows -gt 0) { while ($plan.Count -gt $cap1 + $capN) { $b = Select-DropBand $plan; if (-not $b) { break }; $dropped += , (Remove-Band $plan $b); Remove-EmptyHeads $plan } }
    else { [void]$warn.Add('쪽수를 계산하지 못해 2쪽 맞춤을 건너뛰었습니다.') }
    $pager = @{ used = 0; cap = $(if ($pageRows -gt 0) { $cap1 } else { 0 }); next = $capN; header = $false }
    $row = Write-Plan $ws $sp $plan 7 '경영전략본부' $pager $dates ([ref]$markSeq)
    Complete-Labels $ws $sp 7 ($row - 1); Set-PrintSetup $ws $sp ($row - 1)
    Write-Dropped $ws $sp 7 $dropped
    if ($dropped.Count) { [void]$warn.Add("금주 요약본이 2쪽을 넘어 $($dropped.Count)개 항목을 표 오른쪽 밖(노란 제목 아래)으로 뺐습니다. 확인해 주세요.") }

    [void]$wb.Worksheets.Item(1).Activate()
    $outName = "경영전략본부_주간업무보고_$dateTag.xlsx"; $outPath = Join-Path $outDir $outName
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    [void]$wb.SaveAs($outPath, 51); [void]$wb.Close($false)
    return @{ outPath = $outPath; outName = $outName; warnings = @($warn) }
  } finally { foreach ($w in $open.Values) { try { $w.Close($false) } catch {} } }
}

# ---------- 전사 통합 ----------

function Merge-Company($Excel, $files, [string]$outDir, [string]$dateTag) {
  $warn = New-Object System.Collections.ArrayList
  $open = @{}
  try {
    foreach ($k in $files.Keys) { $open[$k] = $Excel.Workbooks.Open($files[$k], 0, $true) }
    $sum = @{}; foreach ($k in $files.Keys) { $sum[$k] = Find-ReportSheet $open[$k] '요약' }
    $dates = Read-WeekDates $sum['경영전략본부']
    $wb = $Excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { [void]$wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $sp = $script:SheetSpecs.co; $ws = New-ReportSheet $wb $sp $true
    $pageRows = Measure-PageRows $ws $sp
    if ($pageRows -le 0) { [void]$warn.Add('쪽수를 계산하지 못해 쪽 나눔·2쪽 맞춤을 건너뛰었습니다.') }
    $hdrRows = [Math]::Ceiling((16.2 + $sp.titleH + $sp.gapH + $sp.hdrH * 2) / $sp.bodyH)
    $cap = $(if ($pageRows -gt 0) { $pageRows - $hdrRows - 1 } else { 0 })
    $bandSeq = 0; $markSeq = 0; $row = 2; $first = $true

    foreach ($pg in $script:CompanyPages) {
      if (-not $first) { [void]$ws.HPageBreaks.Add($ws.Cells($row, 1)) }
      Write-PageHeader $ws $sp $row $dates; $row += 5; $first = $false
      $plans = @(); $total = 0
      foreach ($blk in $pg.blocks) {
        $srcs = @(); foreach ($part in $blk.parts) { if ($sum[$part]) { $srcs += @{ rs = $sum[$part] } } }
        $plan = New-BlockPlan $srcs ([ref]$bandSeq)
        foreach ($in in @($blk.inserts)) { if ($in -and $sum[$in.part]) { Add-IntoPlan $plan (New-SubRows $sum[$in.part] $in.bracket ([ref]$bandSeq)) $in.after $in.atEnd } }
        $plans += @{ label = $blk.label; plan = $plan }; $total += $plan.Count
      }
      $dropped = @()
      if ($pg.max -gt 0 -and $cap -gt 0) {
        while ($total -gt $cap * $pg.max) {
          $b = Select-DropBand $plans[-1].plan; if (-not $b) { break }
          $dropped += , (Remove-Band $plans[-1].plan $b); Remove-EmptyHeads $plans[-1].plan; $total = ($plans | ForEach-Object { $_.plan.Count } | Measure-Object -Sum).Sum
        }
      }
      $grpStart = $row; $pager = @{ used = 0; cap = $cap; next = $cap; header = $true }
      foreach ($p in $plans) { $row = Write-Plan $ws $sp $p.plan $row $p.label $pager $dates ([ref]$markSeq) }
      Write-Dropped $ws $sp $grpStart $dropped
      if ($dropped.Count) { [void]$warn.Add("경영지원본부가 2쪽을 넘어 $($dropped.Count)개 항목을 표 오른쪽 밖(노란 제목 아래)으로 뺐습니다. 확인해 주세요.") }
    }
    Complete-Labels $ws $sp 2 ($row - 1); Set-PrintSetup $ws $sp ($row - 1)

    # 디지털전략본부: 그 파일의 요약 시트를 그대로 별도 시트로
    if ($sum['디지털전략본부']) {
      [void]$sum['디지털전략본부'].ws.Copy([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
      $wb.Worksheets.Item($wb.Worksheets.Count).Name = '디지털전략본부'
    }
    [void]$wb.Worksheets.Item(1).Activate()
    $outName = "전사_주간업무보고_$dateTag.xlsx"; $outPath = Join-Path $outDir $outName
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    [void]$wb.SaveAs($outPath, 51); [void]$wb.Close($false)
    return @{ outPath = $outPath; outName = $outName; warnings = @($warn) }
  } finally { foreach ($w in $open.Values) { try { $w.Close($false) } catch {} } }
}
