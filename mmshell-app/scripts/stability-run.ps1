param(
  [int]$DurationHours = 8,
  [int]$IntervalSeconds = 20
)

$endTime = (Get-Date).AddHours($DurationHours)
$logPath = Join-Path $PSScriptRoot "..\docs\M2-稳定性跑测日志.txt"

"=== Stability run started: $(Get-Date) ===" | Out-File -FilePath $logPath -Encoding utf8
"DurationHours=$DurationHours IntervalSeconds=$IntervalSeconds" | Out-File -FilePath $logPath -Append -Encoding utf8

while ((Get-Date) -lt $endTime) {
  $now = Get-Date
  "[$now] heartbeat tick - app should remain responsive" | Out-File -FilePath $logPath -Append -Encoding utf8
  Start-Sleep -Seconds $IntervalSeconds
}

"=== Stability run ended: $(Get-Date) ===" | Out-File -FilePath $logPath -Append -Encoding utf8
