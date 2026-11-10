param([Parameter(Mandatory=$true)][string]$Field)
# Puts one card field into the clipboard only. Never prints the value (length only).
$drive = -join ([char[]](0xB9C8,0xBAA8,0xB8E8,0xB4DC,0xB77C,0xC774,0xBE0C))
$path = "C:\$drive\ClaudeSecrets\epost-card.env"
$map = @{}
foreach ($line in [IO.File]::ReadAllLines($path, [Text.Encoding]::UTF8)) {
  if ($line -match '^([A-Z_0-9]+)=(.*)$') { $map[$Matches[1]] = ($Matches[2] -replace '\D','') }
}
$n = $map['CARD_NUMBER']; $pw = $map['CARD_PW2']
switch ($Field) {
  'creditNo1'   { $v = $n.Substring(0,4) }
  'creditNo2'   { $v = $n.Substring(4,4) }
  'creditNo3'   { $v = $n.Substring(8,4) }
  'creditNo4'   { $v = $n.Substring(12,4) }
  'creditExp1'  { $v = $map['CARD_EXP_MM'] }
  'creditExp2'  { $v = $map['CARD_EXP_YY'] }
  'creditPwd1'  { $v = $pw.Substring(0,1) }
  'creditPwd2'  { $v = $pw.Substring(1,1) }
  'creditBirth' { $v = $map['CARD_OWNER_ID'] }
  'clear'       { $v = ' ' }
  default { throw "unknown field" }
}
Set-Clipboard -Value $v
"$Field ready (len $($v.Trim().Length))"
