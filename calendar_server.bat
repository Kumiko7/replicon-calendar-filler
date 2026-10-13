@echo off
cd /d "%~dp0"
echo Calendar link for last week:
powershell -NoProfile -Command "$t=(Get-Date).Date; $w=[int]$t.DayOfWeek; if($w -eq 0){$w=7}; $from=$t.AddDays(-$w-6).ToString('yyyy-MM-dd'); $to=$t.AddDays(-$w).ToString('yyyy-MM-dd'); Write-Host ('http://localhost:8000/calendar?from={0}&to={1}' -f $from, $to)"
echo.
python .\calendar_server.py