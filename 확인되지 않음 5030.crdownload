# 주간업무보고 양식 통일 엔진 (Excel COM 사용)
# server.ps1에서 dot-source 해서 사용: Convert-WeeklyReport -Excel $xl -InPath ... -OutDir ... -OrigName ...

$script:CHK = [char]::ConvertFromUtf32(0x1F5F8)
# 체크 모양: √ ✓ ✔ ☑ ✅ / 🗸 🗹(상자 체크) / ⅴ ｖ Ｖ ∨ (v처럼 생긴 체크)
$script:CheckRx = '(' + ((@($script:CHK, [char]::ConvertFromUtf32(0x1F5F9), '√', '✓', '✔', '☑', '✅', 'ⅴ', 'ｖ', 'Ｖ', '∨') | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')'
# 엑셀 [기호]로 넣은 기호 글꼴의 체크 (셀에는 ü·þ·P·R·a 같은 일반 글자로 저장되고 글꼴 때문에 체크로 보임)
$script:SymbolFontRx = 'Wingdings|Webdings|Symbol|Marlett'
$script:SymbolChecks = @{ 'Wingdings' = 'üþ'; 'Wingdings 2' = 'PRTO'; 'Wingdings 3' = ''; 'Webdings' = 'a' }
$script:FB = "네오현대 B"
$script:FL = "네오현대 L"
$script:DeptSuffix = '(?:본부|담당|팀|실|센터|부문)'
$script:DeptRx = '[가-힣A-Za-z0-9&·]+?' + $script:DeptSuffix + '(?![가-힣])'

function Get-IndentLen([string]$t) {
  $i = 0
  while ($i -lt $t.Length -and ([char]::IsWhiteSpace($t[$i]) -or [int]$t[$i] -in @(0x200B, 0xFEFF, 0x3164))) { $i++ }
  return $i
}

function Parse-Line([string]$raw) {
  $ind = Get-IndentLen $raw
  $s = $raw.Substring($ind)
  $o = @{ raw = $raw; indent = $ind; type = 'none'; body = $s; bodyOff = $ind; textStart = $ind }
  $m = $null
  # 대표 항목 기호: 여러 모양의 체크 표시 + 네모·마름모·삼각형 계열 (□ ■ ▣ ◆ ◇ ▶ ▷ ►)
  # 체크 뒤에 붙는 이모지 표시 문자(U+FE0F/FE0E)도 함께 지운다
  if ($s -match ('^(' + $script:CheckRx + '|□|■|▣|◆|◇|▶|▷|►)[︎️]?\s*')) { $o.type = 'L1'; $m = $matches[0] }
  elseif ($s -match '^\[') { $o.type = 'BR' }
  elseif ($s -match '^※') { $o.type = 'NOTE' }
  elseif ($s -match '^(->|[-_–—−‐])\s*') { if ($matches[1] -eq '->') { $o.type = 'L3s' } else { $o.type = 'L2' }; $m = $matches[0] }
  elseif ($s -match '^(·|•|〮|ㆍ|∙|‧|ㄴ|→|>|\*|▪|◦|○|●|◎|:|\.)\s*') { $o.type = 'L3s'; $m = $matches[0] }
  elseif ($s -match '^(\(\d+\)|\d+\)|[①-⑳]|\d+\.\s)') { $o.type = 'NUM' }
  if ($m) { $o.body = $s.Substring($m.Length); $o.bodyOff = $ind + $m.Length; $o.textStart = $ind + $m.Length }
  elseif ($o.type -eq 'none') { $o.textStart = $ind - 1 }
  $o.body = $o.body.TrimEnd()
  return $o
}

function Get-Fmt($cell, [string]$text) {
  $f = $cell.Font
  $info = @{ size = $f.Size; color = $f.Color; ul = $f.Underline; st = $f.Strikethrough; it = $f.Italic; name = $f.Name; bold = $f.Bold; chars = $null }
  $mixed = ($info.size -is [DBNull]) -or ($info.color -is [DBNull]) -or ($info.ul -is [DBNull]) -or ($info.st -is [DBNull]) -or ($info.it -is [DBNull]) -or ($info.name -is [DBNull]) -or ($info.bold -is [DBNull])
  if ($mixed -and $text.Length -le 2000) {
    $arr = New-Object object[] $text.Length
    for ($i = 1; $i -le $text.Length; $i++) {
      $cf = $cell.Characters($i, 1).Font
      $arr[$i - 1] = @($cf.Size, $cf.Color, $cf.Underline, $cf.Strikethrough, $cf.Italic, $cf.Name, $cf.Bold)
    }
    $info.chars = $arr
    if ($info.size -is [DBNull]) { $info.size = $arr[[int][Math]::Floor(($text.Length - 1) / 2)][0] }
    # 엑셀 색 값은 Double 이어야 한다 (정수 0을 넣으면 '지정한 캐스트가 잘못되었습니다' 오류)
    if ($info.color -is [DBNull]) { $info.color = [double]0 }
  }
  return $info
}

function Test-Val($v) { return ($null -ne $v -and -not ($v -is [DBNull])) }

# ---------- 오탈자 교정 ----------
# 틀린 게 확실한 일반 단어만 고친다. 업계 용어·영어·숫자는 넣지 않는다.
# 글자 수가 같은 교정만 허용해 글자 단위 서식(색·밑줄) 위치가 어긋나지 않게 한다.
$script:TypoRules = @(
  @('(?<![가-힣])맟(?![가-힣])', '및'),
  @('헌황', '현황'),
  @('진황(?=\s*현황)', '진행'),
  @('역활', '역할'),
  @('갯수', '개수'),
  @('됬', '됐'),
  @('되서(?=[\s,.)]|$)', '돼서'),
  @('되요(?=[\s,.)!?]|$)', '돼요'),
  @('몇일', '며칠'),
  @('오랫만', '오랜만'),
  @('일일히', '일일이'),
  @('곰곰히', '곰곰이'),
  @('깨끗히', '깨끗이'),
  @('틈틈히', '틈틈이'),
  @('번번히', '번번이'),
  @('일찍히', '일찍이'),
  @('꼼꼼이', '꼼꼼히'),
  @('솔직이', '솔직히'),
  @('익숙치', '익숙지'),
  @('요컨데', '요컨대'),
  @('생각컨데', '생각건대'),
  @('할려고', '하려고'),
  @('어떻해', '어떻게'),
  @('어짜피', '어차피'),
  @('엇그제', '엊그제'),
  @('왠일', '웬일'),
  @('웬지', '왠지'),
  @('왠만', '웬만'),
  @('희안', '희한'),
  @('어의없', '어이없'),
  # 율/률: 모음·ㄴ 받침 뒤는 '율', 그 밖의 받침 뒤는 '률'
  @('수익율', '수익률'), @('진행율', '진행률'), @('달성율', '달성률'), @('성장율', '성장률'),
  @('합격율', '합격률'), @('가동율', '가동률'), @('이익율', '이익률'), @('투입율', '투입률'),
  @('출석율', '출석률'), @('적중율', '적중률'), @('진척율', '진척률'), @('참석율', '참석률'),
  @('응답율', '응답률'), @('채택율', '채택률'), @('사용율', '사용률'), @('활용율', '활용률'),
  @('불량율', '불량률'), @('점검율', '점검률'), @('수금율', '수금률'),
  @('출하률', '출하율'), @('처리률', '처리율'), @('전환률', '전환율'), @('점유률', '점유율'),
  @('회수률', '회수율'), @('감소률', '감소율'), @('할인률', '할인율'), @('증가률', '증가율'),
  @('원가률', '원가율'), @('접수률', '접수율'), @('완료률', '완료율'), @('준수률', '준수율'),
  @('참여률', '참여율'), @('(?<![가-힣])비률', '비율'),
  @('보급율', '보급률'), @('상승율', '상승률'), @('하락율', '하락률'), @('가입율', '가입률'), @('이직율', '이직률'),
  @('경쟁율', '경쟁률'), @('정확율', '정확률'), @('계약율', '계약률'), @('매출율', '매출률'), @('이행율', '이행률'),
  @('검출율', '검출률'), @('고장율', '고장률'), @('출동율', '출동률'), @('성공율', '성공률'), @('예측율', '예측률'),
  @('절감율', '절감률'), @('감축율', '감축률'), @('인상율', '인상률'), @('공정율', '공정률'), @('충족율', '충족률'),
  @('반영율', '반영률'), @('적용율', '적용률'), @('응찰율', '응찰률'), @('낙찰율', '낙찰률'), @('입찰율', '입찰률'),
  @('수주률', '수주율'), @('갱신률', '갱신율'), @('해지률', '해지율'), @('납부률', '납부율'), @('개선률', '개선율'),
  @('인하률', '인하율'), @('진도률', '진도율'),
  # 자주 틀리는 일반 단어
  @('피트백', '피드백'), @('헷지', '헤지'), @('댓가', '대가'), @('촛점', '초점'), @('싯가', '시가'),
  @('내노라', '내로라'), @('닥달', '닦달'), @('덥썩', '덥석'), @('움추', '움츠'), @('안밖', '안팎'),
  @('오랜동안', '오랫동안'), @('짜집기', '짜깁기'), @('설겆이', '설거지'), @('어줍잖', '어쭙잖'), @('어떡게', '어떻게'),
  @('([가-힣]?[할될줄올갈릴볼알낼울킬])께(?=요|[\s,.)!?]|$)', '$1게'),
  @('([았었였겠했있없])슴(?=[\s,.)/]|$)', '$1음')
)

# 확실하지 않은 것: 고치지 않고 노란색으로 표시만 한다. {0} 자리에 찾은 글자가 들어간다
$script:SuspectRules = @(
  @('(대금|카드|비용|현금|자금|송금|선급금|잔금)\s*결재', '{0} (결재→결제?)'),
  @('결제\s*(상신|라인|권자|서류|문서)', '{0} (결제→결재?)'),
  @('(공시|공지|기사|홈페이지|게시판|사보)\s*개재', '{0} (개재→게재?)'),
  @('(?<![가-힣])일부로(?![가-힣])', '{0} (일부러?)'),
  @('금새', '{0} (금세?)'),
  @('몇\s+일(?![가-힣])', '{0} (며칠?)'),
  @('않(되|돼|됨|된)', '{0} (안 되?)'),
  @('(?<![가-힣])구지(?![가-힣])', '{0} (굳이?)'),
  @('설레임', '{0} (설렘?)'),
  @('되뇌이', '{0} (되뇌?)'),
  @('(?<![가-힣])왠(?=\s)', '{0} (웬?)'),
  @('(?<![0-9가-힣])년차', '{0} (연차?)'),
  @('(?<=[가-힣])(을을|를를|는는|에에서)', '{0} (조사 반복?)'),
  @('(?<![가-힣A-Za-z0-9])([가-힣A-Za-z0-9]{2,})\s+\1(?![가-힣A-Za-z0-9])', '{0} (같은 단어 반복?)')
)

# 자모 한 글자만 다른지 비교할 업무 단어 (맞춤법 검사기가 틀렸다고 한 단어만 비교한다)
$script:Vocab = @'
현황 검토 진행 완료 보고 계획 예정 추진 확인 점검 관리 요청 협의 회의 결과 분석 개선 대응 작성 실시 운영 준비 지원 수립 공유 발송 제출
접수 승인 결재 결제 계약 체결 수주 매출 실적 목표 달성 예산 비용 원가 견적 단가 금액 수익 손익 이익 채권 채무 회수 청구 지급 납부 정산
마감 결산 감사 평가 교육 채용 면접 인사 노무 안전 사고 재해 예방 조치 개정 변경 도입 구축 개발 설계 시스템 프로세스 전략 사업 영업 서비스
설치 공사 현장 고객 협력사 법인 본부 담당 팀장 지사 센터 본사 해외 국내 시장 경쟁 동향 조사 자료 문서 보고서 회의록 일정 기한 기간 인원
조직 업무 과제 이슈 리스크 품질 불량 클레임 교체 수리 보수 갱신 유지 모니터링 피드백 리뷰 업데이트 프로젝트 미팅 세미나 출장 방문 참석 발표
협조 검수 입고 출하 배송 물량 재고 구매 발주 조달 소싱 공급 납기 생산 제작 제조 시험 인증 승강기 엘리베이터 에스컬레이터 정기 수시 긴급 신규
기존 추가 보완 수정 반영 적용 확대 축소 감소 증가 상승 하락 전환 이관 인수 인계 통보 안내 공지 게시 홍보 마케팅 브랜드 행사 전시회 상품
제품 기종 모델 부품 자재 장비 설비 기계 전기 제어 데이터 대시보드 보안 개인정보 규정 지침 기준 절차 양식 서식 표준 가이드 매뉴얼 법률 소송
판결 계약서 약정 협약 합의 이사회 주주 공시 투자 자금 대출 금융 환율 세금 세무 회계 재무 연결 별도 내부 외부 감사인 위원회 경영 대표이사
상무 전무 부장 차장 과장 대리 사원 매니저 직원 임직원 근로자 담당자 책임자 관리자 결정 판단 논의 마무리 시행 착수 착공 준공 개시 종료 연장
만기 재계약 해지 취소 보류 지연 단축 조율 협상 제안 입찰 낙찰 선정 평가표 심사 실사 진단 컨설팅 용역 위탁 수탁 인센티브 성과급 급여 임금
수당 휴가 근태 출근 퇴근 수익률 점유율 달성률 진행률 가동률 출하율 처리율 전환율 회수율 증가율 감소율 원가율 이익률 성장률 계약률
'@ -split '\s+' | Where-Object { $_ } | Sort-Object Length -Descending

# 맞춤법 검사기가 모르는 업계 용어·외래어 (노란색 표시에서 제외)
$script:SpellAllow = @{}
@'
소싱 모듈 쉬브 시브 로프 피트 카도어 카상부 승강로 기계실 권상기 조속기 완충기 이노블록 이노블럭 갇힘 리모델링 모더니제이션 에스컬레이터
무빙워크 주차기 넥스 럭센 헬리오스 챗봇 대시보드 필드웍 컨택 리플릿 룰렛 에코백 직영 대리점 협력사 숙당직 헤지 헷지 온핸드 스터핑 운행량
'@ -split '\s+' | Where-Object { $_ } | ForEach-Object { $script:SpellAllow[$_] = 1 }

# 단어 끝에 붙는 조사·어미 (떼어 내고 단어 본체만 비교한다)
$script:Particles = @('으로써', '으로서', '에서는', '에서', '으로', '까지', '부터', '에게', '이며', '하여', '하고', '했음', '됨', '함', '을', '를', '이', '가', '은', '는', '의', '에', '로', '과', '와', '도', '만', '및', '등')

if (-not ('JamoNear2' -as [type])) {
  Add-Type -TypeDefinition @'
public static class JamoNear2 {
  static int[] J(char c) { int i = c - 0xAC00; if (i < 0 || i > 11171) return new int[] { -1, (int)c, 0 }; return new int[] { i / 588, (i % 588) / 28, i % 28 }; }
  // tok 전체가 vocab 단어 중 하나와 자모 한 개만 다르면 그 단어를 돌려준다. 없으면 null
  public static string Find(string tok, string[] vocab) {
    foreach (string w in vocab) {
      if (w.Length != tok.Length || w == tok) continue;
      int d = 0;
      for (int k = 0; k < w.Length && d < 2; k++) {
        int[] a = J(tok[k]), b = J(w[k]);
        if (a[0] == -1 || b[0] == -1) { if (tok[k] != w[k]) d += 2; continue; }
        if (a[0] != b[0]) d++; if (a[1] != b[1]) d++; if (a[2] != b[2]) d++;
      }
      if (d == 1) return w;
    }
    return null;
  }
}
'@
}
$script:VocabArr = [string[]]$script:Vocab
$script:SpellCache = @{}

# 애매한 오탈자 후보를 찾는다 (고치지 않음). 반환: '헌황→현황?' 같은 설명 목록
function Find-Suspects([string]$s, $app) {
  $res = New-Object System.Collections.ArrayList
  foreach ($rule in $script:SuspectRules) {
    foreach ($m in [regex]::Matches($s, $rule[0])) { [void]$res.Add(($rule[1] -f $m.Value)) }
  }
  foreach ($m in [regex]::Matches($s, '(?<![가-힣A-Za-z0-9])[가-힣]{2,}(?![A-Za-z0-9])')) {
    $tok = $m.Value
    if (-not $script:SpellCache.ContainsKey($tok)) {
      $hit = $null
      # 단어 본체 후보: 그대로 + 조사를 뗀 것
      $stems = @($tok) + @($script:Particles | Where-Object { $tok.EndsWith($_) -and $tok.Length - $_.Length -ge 2 } | ForEach-Object { $tok.Substring(0, $tok.Length - $_.Length) })
      foreach ($st in $stems) {
        if ($script:SpellAllow.ContainsKey($st)) { break }
        $ok = $true; try { $ok = [bool]$app.CheckSpelling($st) } catch {}
        if ($ok) { continue }
        $w = [JamoNear2]::Find($st, $script:VocabArr)
        if ($w) { $hit = "$st→$w"; break }
      }
      $script:SpellCache[$tok] = $hit
    }
    if ($script:SpellCache[$tok]) { [void]$res.Add($script:SpellCache[$tok] + '?') }
  }
  return , @($res | Select-Object -Unique)
}

function Repair-Typos([string]$s, $stats) {
  foreach ($rule in $script:TypoRules) {
    $ms = [regex]::Matches($s, $rule[0])
    if ($ms.Count -eq 0) { continue }
    $new = [regex]::Replace($s, $rule[0], $rule[1])
    if ($new.Length -ne $s.Length) { continue }
    foreach ($m in $ms) { [void]$stats.typos.Add(('{0}→{1}' -f $m.Value, $m.Result($rule[1]))) }
    $s = $new
  }
  return $s
}

# ---------- 긴 문장 줄바꿈 ----------
Add-Type -AssemblyName System.Drawing
$script:Gfx = [System.Drawing.Graphics]::FromImage((New-Object System.Drawing.Bitmap 1, 1))
$script:Gfx.PageUnit = [System.Drawing.GraphicsUnit]::Point
$script:Sf = [System.Drawing.StringFormat]::GenericTypographic.Clone()
$script:Sf.FormatFlags = $script:Sf.FormatFlags -bor [System.Drawing.StringFormatFlags]::MeasureTrailingSpaces
$script:FontCache = @{}
$script:FontAlias = @{ '네오현대 B' = 'NeoHyundai B'; '네오현대 L' = 'NeoHyundai L' }

function Measure-Pt([string]$text, [string]$fontName, [double]$size) {
  $key = "$fontName|$size"
  if (-not $script:FontCache.ContainsKey($key)) {
    $fam = $script:FontAlias[$fontName]; if (-not $fam) { $fam = $fontName }
    $f = New-Object System.Drawing.Font($fam, [single]$size, [System.Drawing.GraphicsUnit]::Point)
    if ($f.Name -ne $fam) { $f = New-Object System.Drawing.Font('Malgun Gothic', [single]$size, [System.Drawing.GraphicsUnit]::Point) }
    $script:FontCache[$key] = $f
  }
  return $script:Gfx.MeasureString($text, $script:FontCache[$key], 100000, $script:Sf).Width
}

# 칸 너비(W, pt)에 맞게 body를 띄어쓰기 단위로 나눈다. 반환: @{ text; start; len } 목록
function Test-BulletStart([string]$s) {
  return ($s -match ('^(' + $script:CheckRx + '|[-_–—−‐·•〮ㆍ∙‧ㄴ→>*▪◦○●◎:.□■▣◆◇▶▷►※]|[①-⑳]|\(\d+\)|\d+\)|\d+\.\s)'))
}

# 줄 첫 기호가 기호 글꼴(Wingdings 등)로 된 글자면 실제 모양에 맞게 다시 분류한다.
# 체크 모양이면 대표 항목(L1), 그 밖의 기호(●, ■ 등)는 하위 항목 기호로 본다.
function Resolve-SymbolBullet($p) {
  # 기호 글꼴의 글자는 일반 글자(ü, P, a …)나 0xF000대 글자라서 '기호 없는 줄'로 먼저 분류된다
  if ($p.type -ne 'none' -or $p.indent -ge $p.raw.Length) { return }
  $ch = $p.raw[$p.indent]
  if ([char]::IsWhiteSpace($ch) -or ([int]$ch -ge 0xAC00 -and [int]$ch -le 0xD7A3)) { return }
  $fn = ''; try { $fn = [string]$p.cell.Characters($p.indent + 1, 1).Font.Name } catch {}
  if ($fn -notmatch $script:SymbolFontRx) { return }
  $isCheck = $false
  foreach ($k in $script:SymbolChecks.Keys) { if ($fn -eq $k -and $script:SymbolChecks[$k].Contains([string]$ch)) { $isCheck = $true } }
  # 기호 글꼴의 글자는 0xF000대(PUA)로 들어오는 경우도 있다: F0FC=ü(체크), F0FE=þ(상자 체크)
  if ([int]$ch -in @(0xF0FC, 0xF0FE)) { $isCheck = $true }
  $len = 1; while ($p.indent + $len -lt $p.raw.Length -and [char]::IsWhiteSpace($p.raw[$p.indent + $len])) { $len++ }
  $p.type = $(if ($isCheck) { 'L1' } else { 'L3s' })
  $p.body = $p.raw.Substring($p.indent + $len).TrimEnd(); $p.bodyOff = $p.indent + $len; $p.textStart = $p.indent + $len
  $p.symbolMk = 'SYM:' + $fn + [string]$ch
}

function Split-ToWidth([string]$prefix, [string]$body, [int]$pad, [string]$fontName, [double]$size, [double]$W) {
  $res = @(); $pos = 0; $first = $true
  while ($pos -lt $body.Length -and $res.Count -lt 50) {
    $lead = $(if ($first) { $prefix } else { ' ' * $pad })
    $rest = $body.Substring($pos)
    if ((Measure-Pt ($lead + $rest) $fontName $size) -le $W) { $res += @{ text = $lead + $rest; start = $pos; len = $rest.Length }; break }
    # 한 문장에는 글머리 기호 하나만: 내린 줄이 '- 계속'처럼 기호로 시작하게 되는 자리에서는 끊지 않는다
    $cut = -1
    for ($i = 1; $i -lt $rest.Length; $i++) {
      if ($rest[$i] -ne ' ') { continue }
      if ((Measure-Pt ($lead + $rest.Substring(0, $i).TrimEnd()) $fontName $size) -gt $W) { break }
      if (-not (Test-BulletStart $rest.Substring($i).TrimStart())) { $cut = $i }
    }
    if ($cut -le 0) {
      # 띄어쓰기로 못 나누면 글자 단위로 자른다 (이때도 다음 줄이 기호로 시작하지 않게)
      $cut = 1
      while ($cut -lt $rest.Length -and (Measure-Pt ($lead + $rest.Substring(0, $cut + 1)) $fontName $size) -le $W) { $cut++ }
      while ($cut -gt 1 -and (Test-BulletStart $rest.Substring($cut).TrimStart())) { $cut-- }
    }
    $piece = $rest.Substring(0, $cut).TrimEnd()
    $res += @{ text = $lead + $piece; start = $pos; len = $piece.Length }
    $pos += $cut
    while ($pos -lt $body.Length -and $body[$pos] -eq ' ') { $pos++ }
    $first = $false
  }
  return , $res
}

# 엑셀이 글자 서식 값을 비워서(DBNull) 주는 경우가 있어, 값이 있을 때만 넣고 실패해도 변환은 계속한다
function Set-FontProp($font, [string]$prop, $val) {
  if (-not (Test-Val $val)) { return }
  if ($prop -in @('Color', 'Size') -and $val -is [int]) { $val = [double]$val }
  try { $font.$prop = $val } catch {}
}

# 내용 칸 머리글('전주·금주' 또는 '금주·차주')을 찾아 칸 배치를 알아낸다.
# 머리글 이름은 칸 위치를 찾는 데만 쓰고, 양식 통일 방식은 어느 쪽이든 같다.
$script:HeadLeft = @('전주', '금주')
$script:HeadRight = @{ '전주' = @('금주'); '금주' = @('차주', '다음주') }

function Find-Layout($ws) {
  $rng = $ws.Range($ws.Cells(1, 1), $ws.Cells(25, 26))
  $vals = $rng.Value2
  $best = $null
  for ($r = 1; $r -le 25; $r++) {
    for ($c = 2; $c -le 25; $c++) {
      $t = ([string]$vals[$r, $c]) -replace '\s', ''
      $left = $script:HeadLeft | Where-Object { $t.StartsWith($_) } | Select-Object -First 1
      if (-not $left) { continue }
      for ($c2 = $c + 1; $c2 -le [Math]::Min(26, $c + 2); $c2++) {
        $t2 = ([string]$vals[$r, $c2]) -replace '\s', ''
        if (@($script:HeadRight[$left] | Where-Object { $t2.StartsWith($_) }).Count -gt 0) {
          if ($null -eq $best -or $c -lt $best.prev) { $best = @{ row = $r; prev = $c; cur = $c2 } }
          break
        }
      }
    }
  }
  if ($null -eq $best) { return $null }
  $start = $best.row + 1
  while ($start -le $best.row + 4 -and ([string]$ws.Cells($start, $best.prev).Value2) -match '^\s*\(?\s*\d{1,2}\s*/\s*\d{1,2}') { $start++ }
  return @{ headerRow = $best.row; lab = $best.prev - 1; cols = @($best.prev, $best.cur); start = $start }
}

function Clean-Label([string]$v) {
  $lines = @($v -split "[`r`n]+" | ForEach-Object { ($_ -replace '\s', '') } | Where-Object { $_ -ne '' })
  $parts = @(); $buf = ''
  foreach ($l in $lines) {
    $buf += $l
    if ($buf -match ($script:DeptSuffix + '$')) { $parts += $buf; $buf = '' }
  }
  if ($buf) { $parts += $buf }
  return ($parts -join ',')
}

function Get-Sections($ws, $lay) {
  # 숨긴 행이 있으면 End(xlUp)가 마지막 행을 못 찾을 수 있어, 사용 범위 끝까지 실제 값으로 확인한다
  $lastRow = 0
  $urEnd = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  foreach ($col in $lay.cols) {
    $e = $ws.Cells($ws.Rows.Count, $col).End(-4162).Row
    if ($urEnd -gt $e) {
      $vals = $ws.Range($ws.Cells($e, $col), $ws.Cells($urEnd, $col)).Value2
      if ($vals -is [array]) { for ($i = $vals.GetLength(0); $i -ge 1; $i--) { if (([string]$vals[$i, 1]).Trim() -ne '') { $e = $e + $i - 1; break } } }
    }
    if ($e -gt $lastRow) { $lastRow = $e }
  }
  $starts = @(); $labels = @()
  $r = $lay.start
  while ($r -le $lastRow) {
    $lc = $ws.Cells($r, $lay.lab); $ma = $lc.MergeArea
    $v = [string]$lc.Value2
    if ($v.Trim() -ne '' -or $r -eq $lay.start) { $starts += $r; $labels += $v }
    $r = $ma.Row + $ma.Rows.Count
  }
  $sections = @()
  for ($i = 0; $i -lt $starts.Count; $i++) {
    $s0 = $starts[$i]
    if ($i -lt $starts.Count - 1) { $s1 = $starts[$i + 1] - 1 }
    else { $ma = $ws.Cells($s0, $lay.lab).MergeArea; $s1 = [Math]::Max($lastRow, $ma.Row + $ma.Rows.Count - 1) }
    $sections += , @($s0, $s1, $labels[$i])
  }
  return , $sections
}

# 원본의 부서명 추정: 제목 괄호 → 구분 칸(1개일 때) → 파일명 → 파일명 첫 단어
function Get-DeptName($ws, $lay, $sections, [string]$fileName) {
  $base = [IO.Path]::GetFileNameWithoutExtension($fileName)
  for ($r = 1; $r -lt $lay.headerRow; $r++) {
    for ($c = 1; $c -le 12; $c++) {
      $t = [string]$ws.Cells($r, $c).Value2
      foreach ($pm in [regex]::Matches($t, '\(([^)]*)\)')) {
        $dm = [regex]::Match($pm.Groups[1].Value, $script:DeptRx)
        if ($dm.Success) { return @{ name = $dm.Value; how = '제목' } }
      }
    }
  }
  $labs = @($sections | ForEach-Object { Clean-Label ([string]$_[2]) } | Where-Object { $_ -ne '' } | Select-Object -Unique)
  $deptLabs = @($labs | Where-Object { $_ -match ('^([가-힣A-Za-z0-9&·]+' + $script:DeptSuffix + ',?)+$') })
  if ($labs.Count -eq 1 -and $deptLabs.Count -eq 1) { return @{ name = $deptLabs[0]; how = '구분 칸' } }
  $fm = @([regex]::Matches($base, $script:DeptRx) | ForEach-Object { $_.Value })
  if ($fm.Count -gt 0) { return @{ name = $fm[-1]; how = '파일명' } }
  if ($deptLabs.Count -gt 0) { return @{ name = $deptLabs[0]; how = '구분 칸' } }
  $stop = '^(대표이사|주간업무보고자료|주간업무보고|주간업무|주간보고|주간회의|전사|CSO|R\d+|v\d+|\d+.*|W\d+.*|.*월|.*주차)$'
  $words = @(($base -replace "[()\[\]'_,.\-~]", ' ') -split '\s+' | Where-Object { $_ -ne '' -and $_ -notmatch $stop })
  if ($words.Count -gt 0) { return @{ name = $words[0]; how = '파일명 추정' } }
  return @{ name = '부서명'; how = '찾지 못함' }
}

$script:ContentColWidth = 85

function Process-Sheet($ws, $lay, $stats) {
  # 내용 칸(전주/금주 또는 금주/차주) 열 너비를 85로 통일. 사진·도형이 칸을 따라 늘어나지 않게 크기는 되돌린다
  $needWidth = @($lay.cols | Where-Object { [Math]::Abs($ws.Columns($_).ColumnWidth - $script:ContentColWidth) -gt 0.01 })
  if ($needWidth.Count -gt 0) {
    $sizes = @(); foreach ($sh in $ws.Shapes) { try { $sizes += @{ s = $sh; w = $sh.Width; h = $sh.Height } } catch {} }
    foreach ($col in $needWidth) { $ws.Columns($col).ColumnWidth = [double]$script:ContentColWidth }
    foreach ($z in $sizes) {
      try {
        if ($z.s.Width -ne $z.w -or $z.s.Height -ne $z.h) {
          $lock = $z.s.LockAspectRatio; $z.s.LockAspectRatio = 0
          $z.s.Width = $z.w; $z.s.Height = $z.h
          $z.s.LockAspectRatio = $lock
        }
      } catch {}
    }
    $stats.widened += $needWidth.Count
  }
  $sections = Get-Sections $ws $lay
  $pics = @()
  foreach ($sh in $ws.Shapes) {
    try {
      $tl = $sh.TopLeftCell; $br = $sh.BottomRightCell
      if ($lay.cols -contains $tl.Column) { $pics += @{ shape = $sh; col = $tl.Column; r0 = $tl.Row; r1 = $br.Row; dy = ($sh.Top - $tl.Top) } }
    } catch {}
  }

  for ($si = $sections.Count - 1; $si -ge 0; $si--) {
    $s0 = $sections[$si][0]; $s1 = $sections[$si][1]
    $plans = @()
    foreach ($col in $lay.cols) {
      $vals = $ws.Range($ws.Cells($s0, $col), $ws.Cells($s1, $col)).Value2
      $items = @(); $anchor = 0; $lastContent = 0
      $colPics = @($pics | Where-Object { $_.col -eq $col -and $_.r0 -ge $s0 -and $_.r0 -le $s1 })
      $r = $s0
      while ($r -le $s1) {
        $pic = $colPics | Where-Object { $_.r0 -eq $r } | Select-Object -First 1
        if ($pic) {
          $items += @{ type = 'PIC'; pic = $pic; n = ($pic.r1 - $pic.r0 + 1); gap = $(if ($anchor) { $r - $lastContent - 1 } else { 0 }) }
          if (-not $anchor) { $anchor = $r }
          $lastContent = $pic.r1; $r = $pic.r1 + 1; continue
        }
        if ($s1 -gt $s0) { $v = $vals[($r - $s0 + 1), 1] } else { $v = $vals }
        $t = [string]$v
        if ($t.Trim() -ne '') {
          $p = Parse-Line $t; $p.row = $r; $p.cell = $ws.Cells($r, $col)
          Resolve-SymbolBullet $p
          $p.gap = $(if ($anchor) { $r - $lastContent - 1 } else { 0 })
          if (-not $anchor) { $anchor = $r }
          $lastContent = $r
          $items += $p
        }
        $r++
      }
      if (-not $anchor) { continue }

      # 계층 판별
      # l2Mk: 2단계를 차지한 기호 종류. 다른 기호(•, ① 등)가 2단계면 그보다 깊은 '-'는 3단계로 본다
      $hasL2 = $false; $l2Ind = 0; $l2Mk = ''; $prevLvl = $null; $prevTS = 0; $prevBase = $null
      $hasL3 = $false; $l3Mk = ''; $l3Ind = 0; $prevRawEnd = ''; $prevNoneSub = $false; $prevNoneInd = -1
      foreach ($it in $items) {
        if ($it.type -eq 'PIC') { continue }
        $mk = switch ($it.type) { 'L2' { '-' } 'L3s' { $it.raw.Substring($it.indent, 1) } 'NUM' { 'NUM' } default { $it.type } }
        switch ($it.type) {
          'L1'   { $lvl = 'L1'; $hasL2 = $false }
          'BR'   { $lvl = 'BR'; $hasL2 = $false }
          'NOTE' { $lvl = 'NOTE'; $hasL2 = $false }
          'L2'   {
            if ($hasL2 -and $l2Mk -ne '-' -and $it.indent -gt $l2Ind) {
              if ($hasL3 -and $l3Mk -ne '-' -and $it.indent -gt $l3Ind) { $lvl = 'L4' }
              else { $lvl = 'L3'; $hasL3 = $true; $l3Mk = '-'; $l3Ind = $it.indent }
            } else { $lvl = 'L2' }
          }
          'L3s'  {
            if ($hasL2 -and ($mk -eq $l2Mk -or $it.indent -lt $l2Ind)) { $lvl = 'L2' }
            elseif ($hasL3 -and $mk -ne $l3Mk -and $it.indent -gt $l3Ind) { $lvl = 'L4' }
            elseif ($hasL2) { $lvl = 'L3'; $hasL3 = $true; $l3Mk = $mk; $l3Ind = $it.indent }
            else { $lvl = 'L2' }
          }
          'NUM'  {
            if ($hasL2 -and $l2Mk -eq 'NUM') { $lvl = 'L2' }
            elseif ($hasL2 -and $it.indent -gt $l2Ind) { $lvl = 'L3' } else { $lvl = 'L2' }
          }
          'none' {
            # 기호 없는 줄은 기본적으로 윗줄에 이어지는 문장(글머리 기호 없음)으로 본다.
            # 하위 항목으로 보는 건 윗줄이 ':'로 끝날 때(예: '- 집중시장 :' 아래 '북미 : …')와 그 형제 줄뿐이다.
            $subOk = ($prevRawEnd -match '[:：]$') -or ($prevNoneSub -and $it.indent -eq $prevNoneInd)
            if ($prevLvl -and $prevLvl -notin @('BR', 'NOTE') -and ($it.indent -le $prevTS -or -not $subOk)) { $lvl = 'CONT' }
            elseif ($hasL2) { $lvl = 'L3' } elseif ($prevLvl) { $lvl = 'L2' } else { $lvl = 'RAW' }
          }
        }
        $prevRawEnd = $it.raw.TrimEnd()
        $prevNoneSub = ($it.type -eq 'none' -and $lvl -in @('L2', 'L3')); $prevNoneInd = $it.indent
        # 본문 각 글자가 원본 셀의 몇 번째 글자인지 (글자 단위 색·밑줄을 옮길 때 사용)
        $it.bmap = New-Object int[] $it.body.Length
        for ($q = 0; $q -lt $it.body.Length; $q++) { $it.bmap[$q] = $it.bodyOff + $q }
        # '🗸 [xxx] 내용' → 체크는 남기고 대괄호만 지운다: '🗸 xxx 내용'
        if ($lvl -eq 'L1') {
          $bm = [regex]::Match($it.body, '^\[([^\]]*)\]\s*')
          if ($bm.Success) {
            $inner = $bm.Groups[1].Value.Trim(); $innerAt = $bm.Groups[1].Index + $bm.Groups[1].Value.IndexOf($inner)
            $rest = $it.body.Substring($bm.Length)
            $newMap = New-Object System.Collections.Generic.List[int]
            for ($q = 0; $q -lt $inner.Length; $q++) { $newMap.Add($it.bodyOff + $innerAt + $q) }
            $newBody = $inner
            if ($rest.Length -gt 0) {
              $newBody += ' ' + $rest; $newMap.Add($it.bodyOff + $bm.Groups[1].Index + $bm.Groups[1].Length)
              for ($q = 0; $q -lt $rest.Length; $q++) { $newMap.Add($it.bodyOff + $bm.Length + $q) }
            }
            $it.body = $newBody; $it.bmap = $newMap.ToArray(); $it.debracket = $true
          }
        }
        if ($lvl -eq 'L2') { $hasL2 = $true; $l2Ind = $it.indent; $l2Mk = $mk; $hasL3 = $false }
        if ($lvl -in @('L1', 'L1BR', 'BR', 'NOTE')) { $hasL3 = $false }
        if ($lvl -eq 'CONT') { $it.base = $(if ($prevLvl -eq 'CONT') { $prevBase } else { $prevLvl }); $prevBase = $it.base } else { $prevBase = $lvl }
        $it.lvl = $lvl
        if ($lvl -ne 'CONT') { $prevLvl = $lvl; $prevTS = $it.textStart }
      }

      # 출력 줄 구성 (그룹 사이 빈 줄 최소 1줄, 기존 간격은 유지해 전주/금주 정렬 보존)
      $out = @(); $prevOut = $null
      foreach ($it in $items) {
        $starter = ($it.type -ne 'PIC') -and ($it.lvl -in @('L1', 'L1BR', 'BR', 'NOTE'))
        $nb = 0
        if ($out.Count -gt 0) {
          if ($prevOut -eq 'BR') { $nb = $it.gap }
          elseif ($it.type -eq 'PIC' -or $prevOut -eq 'PIC') { $nb = $it.gap }
          elseif ($starter -or $prevOut -eq 'NOTE') { $nb = [Math]::Max(1, $it.gap) }
        }
        for ($q = 0; $q -lt $nb; $q++) { $out += @{ blank = $true } }
        if ($it.type -eq 'PIC') { $out += @{ pic = $it.pic; n = $it.n }; $prevOut = 'PIC'; continue }
        $wrapable = $it.lvl -in @('L1', 'L1BR', 'L2', 'L3', 'L4', 'CONT')
        if ($wrapable) {
          $it.body = Repair-Typos $it.body $stats
          $it.suspects = Find-Suspects $it.body $ws.Application
          foreach ($x in $it.suspects) { [void]$stats.suspects.Add($x) }
        }
        switch ($it.lvl) {
          'L1'   { $txt = " $($script:CHK) " + $it.body; $font = $script:FB; $boff = 4 }
          'L1BR' { $txt = " " + $it.body; $font = $script:FB; $boff = 1 }
          'L2'   { $txt = "   - " + $it.body; $font = $script:FL; $boff = 5 }
          'L3'   { $txt = "      · " + $it.body; $font = $script:FL; $boff = 8 }
          'L4'   { $txt = "         · " + $it.body; $font = $script:FL; $boff = 11 }
          'CONT' {
            $pad = @{ L1 = 3; L1BR = 1; L2 = 5; L3 = 8; L4 = 11; RAW = 0 }[$it.base]; if ($null -eq $pad) { $pad = 0 }
            $txt = (' ' * $pad) + $it.body
            $font = $(if ($it.base -in @('L1', 'L1BR')) { $script:FB } else { $script:FL }); $boff = $pad
          }
          default { $txt = $it.raw; $font = $null; $boff = 0 }
        }
        # 네오현대에 없는 가운뎃점(∙ U+2219)은 있는 모양(· U+00B7)으로 바꾼다 (글자 수 같음)
        $txt = $txt.Replace([string][char]0x2219, [string][char]0x00B7)
        if ($wrapable) { $it.body = $it.body.Replace([string][char]0x2219, [string][char]0x00B7) }
        $pieces = $null
        if ($wrapable -and -not $it.cell.WrapText) {
          # 칸 너비를 넘으면 아랫줄로 내린다 (윗줄 글자 시작 위치에 맞춰 들여쓰기)
          $W = $(if ($it.cell.MergeCells) { $it.cell.MergeArea.Width } else { $it.cell.Width }) - 3
          $sz = $it.cell.Font.Size; if (-not (Test-Val $sz)) { $sz = $it.cell.Characters([Math]::Min($it.raw.Length, $it.bodyOff + 1), 1).Font.Size }
          if (-not (Test-Val $sz)) { $sz = 11 }
          $pad = @{ L1 = 3; L1BR = 1; L2 = 5; L3 = 8; L4 = 11 }[$it.lvl]; if ($it.lvl -eq 'CONT') { $pad = $boff }
          if ($W -gt 20 -and (Measure-Pt $txt $font $sz) -gt $W) {
            $pieces = Split-ToWidth $txt.Substring(0, $boff) $it.body $pad $font $sz $W
            if ($pieces.Count -le 1) { $pieces = $null }
          }
        }
        if ($pieces) {
          $stats.wrapped++
          for ($pi = 0; $pi -lt $pieces.Count; $pi++) {
            $pc = $pieces[$pi]
            $out += @{ txt = $pc.text; font = $font; boff = $(if ($pi -eq 0) { $boff } else { $pad }); src = $it
                       lvl = $(if ($pi -eq 0) { $it.lvl } else { 'WRAP' }); pstart = $pc.start; plen = $pc.len }
          }
        } else {
          $out += @{ txt = $txt; font = $font; boff = $boff; src = $it; lvl = $it.lvl; pstart = 0; plen = $it.body.Length }
        }
        $prevOut = $it.lvl
        if ($stats.levels.ContainsKey($it.lvl)) { $stats.levels[$it.lvl]++ } else { $stats.levels[$it.lvl] = 1 }
      }
      $need = 0
      foreach ($o in $out) { if ($o.pic) { $need += $o.n } else { $need++ } }
      $plans += @{ col = $col; anchor = $anchor; last = $lastContent; out = $out; need = $need }
    }
    if ($plans.Count -eq 0) { continue }

    $endNeeded = 0; $maxLast = 0
    foreach ($p in $plans) { $e = $p.anchor + $p.need - 1; if ($e -gt $endNeeded) { $endNeeded = $e }; if ($p.last -gt $maxLast) { $maxLast = $p.last } }
    $extra = $endNeeded - $s1

    # 쓰기 전에 원본 서식을 모두 읽어둔다
    foreach ($p in $plans) { foreach ($o in $p.out) { if ($o.src) { if (-not $o.src.fmtc) { $o.src.fmtc = Get-Fmt $o.src.cell $o.src.raw }; $o.fmt = $o.src.fmtc } } }
    if ($extra -gt 0) {
      $ins = [Math]::Min($maxLast + 1, $s1)
      if ($ins -le $maxLast) { $ins = $s1 }
      $ws.Range($ws.Rows($ins), $ws.Rows($ins + $extra - 1)).EntireRow.Insert(-4121, 0) | Out-Null
      $s1 += $extra
      $stats.inserted += $extra
    }
    foreach ($p in $plans) {
      $col = $p.col
      # 행을 기존 내용 사이에 끼워 넣었으면 그 아래 원래 줄이 밀려 내려가 있으므로 그만큼 더 지운다
      $lastNow = $p.last; if ($extra -gt 0 -and $ins -le $p.last) { $lastNow += $extra }
      $clearTo = [Math]::Max($lastNow, $p.anchor + $p.need - 1)
      for ($r = $p.anchor; $r -le $clearTo; $r++) { $cl = $ws.Cells($r, $col); if (-not $cl.MergeCells) { [void]$cl.ClearContents() } }
      $r = $p.anchor
      foreach ($o in $p.out) {
        if ($o.blank) { $r++; continue }
        if ($o.pic) { $o.pic.shape.Top = $ws.Cells($r, $col).Top + $o.pic.dy; $r += $o.n; continue }
        $cell = $ws.Cells($r, $col); $fmt = $o.fmt
        $cell.Value2 = $o.txt
        if ($o.font) {
          $cf = $cell.Font; $cf.Name = $o.font; $cf.Bold = $false
          Set-FontProp $cf 'Size' $fmt.size
          Set-FontProp $cf 'Color' $fmt.color
          $cf.Italic = $false; $cf.Strikethrough = $false; $cf.Underline = -4142
          if ($fmt.chars) {
            $len = $o.plen
            for ($k = 0; $k -lt $len; $k++) {
              if ($o.pstart + $k -ge $o.src.bmap.Length) { break }
              $si2 = $o.src.bmap[$o.pstart + $k]
              if ($si2 -ge $fmt.chars.Count) { break }
              $ch = $fmt.chars[$si2]
              $ul = $(if ((Test-Val $ch[2]) -and $ch[2] -ne -4142) { $ch[2] } else { $null })
              $stk = (Test-Val $ch[3]) -and $ch[3]
              $itl = (Test-Val $ch[4]) -and $ch[4]
              if (((Test-Val $ch[1]) -and $ch[1] -ne $fmt.color) -or $ul -or $stk -or $itl) {
                $tf = $cell.Characters($o.boff + $k + 1, 1).Font
                Set-FontProp $tf 'Color' $ch[1]
                if ($ul) { Set-FontProp $tf 'Underline' $ul }
                if ($stk) { Set-FontProp $tf 'Strikethrough' $true }
                if ($itl) { Set-FontProp $tf 'Italic' $true }
              }
            }
          }
          if ($o.lvl -eq 'L1') { try { $cell.Characters(2, 2).Font.Name = 'Segoe UI Symbol' } catch {} }
          # 오탈자인지 확실하지 않은 줄: 노란색 형광펜 + 메모 (내린 줄까지 모두 칠한다)
          if ($o.src.suspects -and $o.src.suspects.Count -gt 0) {
            try { $cell.Interior.Color = [double]65535 } catch {}
            if ($o.lvl -ne 'WRAP') {
              $note = '오탈자 확인 필요: ' + ($o.src.suspects -join ', ')
              try {
                if ($cell.Comment) { [void]$cell.Comment.Text(([string]$cell.Comment.Text() + "`n" + $note)) }
                else { [void]$cell.AddComment($note) }
              } catch {}
            }
          }
        } else {
          # 대괄호 제목 / ※ 메모: 글자는 그대로 두고 글꼴만 통일 (대괄호 제목 = 네오현대 B, 그 밖 = 네오현대 L).
          # 색·밑줄 등은 원래대로 유지한다
          $uni = $(if ($o.lvl -eq 'BR') { $script:FB } else { $script:FL })
          $cf = $cell.Font; $cf.Name = $uni; $cf.Bold = $false
          Set-FontProp $cf 'Size' $fmt.size; Set-FontProp $cf 'Color' $fmt.color
          Set-FontProp $cf 'Italic' $fmt.it; Set-FontProp $cf 'Underline' $fmt.ul; Set-FontProp $cf 'Strikethrough' $fmt.st
          if ($fmt.chars) {
            for ($k = 0; $k -lt $o.txt.Length -and $k -lt $fmt.chars.Count; $k++) {
              $ch = $fmt.chars[$k]
              if (($ch[1] -ne $fmt.color) -or ($ch[0] -ne $fmt.size) -or ($ch[2] -ne -4142) -or $ch[3] -or $ch[4]) {
                $tf = $cell.Characters($k + 1, 1).Font
                Set-FontProp $tf 'Size' $ch[0]; Set-FontProp $tf 'Color' $ch[1]; Set-FontProp $tf 'Underline' $ch[2]
                Set-FontProp $tf 'Strikethrough' $ch[3]; Set-FontProp $tf 'Italic' $ch[4]
              }
            }
          }
        }
        $r++
      }
    }
  }
}

# 반환: @{ dept; deptHow; outPath; sheets; stats; warnings }
function Convert-WeeklyReport($Excel, [string]$InPath, [string]$OutDir, [string]$OrigName, [string]$DateTag) {
  $warn = New-Object System.Collections.ArrayList
  $wb = $Excel.Workbooks.Open($InPath, 0, $false)
  try {
    $cands = @()
    foreach ($ws in $wb.Worksheets) {
      if ($ws.Visible -ne -1) { continue }
      $lay = Find-Layout $ws
      if ($lay) { $cands += @{ ws = $ws; lay = $lay } }
    }
    if ($cands.Count -eq 0) { throw "'전주 / 금주' 또는 '금주 / 차주' 머리글이 있는 시트를 찾지 못했습니다." }
    if ($cands.Count -gt 3) {
      $act = $wb.ActiveSheet
      $pick = @($cands | Where-Object { $_.ws.Name -eq $act.Name })
      if ($pick.Count -eq 0) { $pick = @($cands[-1]) }
      [void]$warn.Add("보고 시트가 $($cands.Count)개라서 '$($pick[0].ws.Name)' 시트만 바꿨습니다.")
      $cands = $pick
    } else {
      $act = $wb.ActiveSheet
      $cands = @($cands | Sort-Object { if ($_.ws.Name -eq $act.Name) { 0 } else { 1 } })
    }
    $stats = @{ levels = @{}; inserted = 0; wrapped = 0; widened = 0; typos = (New-Object System.Collections.ArrayList); suspects = (New-Object System.Collections.ArrayList) }
    $deptInfo = $null; $sheetNames = @()
    foreach ($c in $cands) {
      if ($null -eq $deptInfo) { $deptInfo = Get-DeptName $c.ws $c.lay (Get-Sections $c.ws $c.lay) $OrigName }
      Process-Sheet $c.ws $c.lay $stats | Out-Null
      $sheetNames += $c.ws.Name
    }
    if ($deptInfo.how -in @('파일명 추정', '찾지 못함')) { [void]$warn.Add("파일 안에서 부서명을 찾지 못해 '$($deptInfo.name)'(으)로 정했습니다. 이름을 확인해 주세요.") }
    if ($stats.levels['L4']) { [void]$warn.Add("4단계 하위 항목 $($stats.levels['L4'])개는 · 기호에 들여쓰기 9칸으로 표시했습니다.") }
    if ($stats.wrapped -gt 0) { [void]$warn.Add("칸을 넘는 긴 문장 $($stats.wrapped)개를 아랫줄로 내렸습니다.") }
    if ($stats.inserted -gt 0) { [void]$warn.Add("빈 줄·줄바꿈 때문에 행 $($stats.inserted)개를 추가했습니다.") }
    # 오탈자 수정 내역: 같은 교정은 묶어서 '맟→및 (2건)' 형태로
    $fixes = @($stats.typos | Group-Object | ForEach-Object { if ($_.Count -gt 1) { "$($_.Name) ($($_.Count)건)" } else { $_.Name } })
    $suspects = @($stats.suspects | Select-Object -Unique)
    if ($stats.widened -gt 0) { [void]$warn.Add("내용 칸 열 너비를 85로 맞췄습니다.") }

    $outName = "{0}_주간업무보고_{1}.xlsx" -f $deptInfo.name, $DateTag
    $outPath = Join-Path $OutDir $outName
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    $wb.SaveAs($outPath, 51)
    return @{ dept = $deptInfo.name; deptHow = $deptInfo.how; outPath = $outPath; outName = $outName; sheets = $sheetNames; stats = $stats; warnings = @($warn); fixes = $fixes; suspects = $suspects }
  } finally {
    $wb.Close($false)
  }
}
